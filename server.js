require('dotenv').config();
const express = require('express');
const axios = require('axios');
const cors = require('cors');
const path = require('path');
const xlsx = require('xlsx');
const session = require('express-session');
const bodyParser = require('body-parser');

const app = express();

app.use(cors());
app.use(bodyParser.json({ limit: '10mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '10mb' }));

app.use(session({
  secret: 'super_secure_secret_change_this',
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: false,
    httpOnly: true,
    maxAge: 60 * 60 * 1500
  }
}));

app.use('/public', express.static(path.join(__dirname, 'public')));
app.get('/login.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});

const {
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  ONEDRIVE_USER,
  ONEDRIVE_FOLDER_PATH,
  LOGIN_USERNAME,
  LOGIN_PASSWORD
} = process.env;

const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';

function requireLogin(req, res, next) {
  if (req.session && req.session.loggedIn) {
    next();
  } else {
    if (req.originalUrl.startsWith('/api/')) {
      return res.status(401).json({ error: 'Unauthorized, please login.' });
    } else {
      return res.redirect(`/login.html?returnTo=${encodeURIComponent(req.originalUrl)}`);
    }
  }
}

app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  if (username === LOGIN_USERNAME && password === LOGIN_PASSWORD) {
    req.session.loggedIn = true;
    req.session.save(() => {
      res.json({ 
        success: true,
        returnTo: req.query.returnTo || '/'
      });
    });
  } else {
    res.status(401).json({ success: false, message: 'Invalid credentials' });
  }
});

app.get('/logout', (req, res) => {
  req.session.destroy(() => {
    res.redirect('/login.html');
  });
});

app.get('/', requireLogin, (req, res) => {
  res.sendFile(path.join(__dirname, 'views', 'index.html'));
});

app.get('/api/list-folders', requireLogin, async (req, res) => {
  const folderPath = req.query.path || ONEDRIVE_FOLDER_PATH;
  try {
    const token = await getAccessToken();
    const url = `${GRAPH_ROOT}/users/${ONEDRIVE_USER}/drive/root:/${folderPath}:/children`;
    const response = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` }
    });
    const folders = response.data.value
      .filter(item => item.folder)
      .map(folder => folder.name);
    res.json(folders);
  } catch (err) {
    console.error('Error listing folders:', err.response?.data || err.message);
    res.status(500).json({ error: 'Failed to list folders', details: err.message });
  }
});

app.get('/api/excel-data', requireLogin, async (req, res) => {
  const { folderName } = req.query;
  if (!folderName) return res.status(400).json({ error: 'Missing folderName' });

  try {
    const token = await getAccessToken();

    // First, list all files in the folder to find the most recent Excel file
    const folderUrl = `${GRAPH_ROOT}/users/${ONEDRIVE_USER}/drive/root:/${ONEDRIVE_FOLDER_PATH}/${folderName}:/children`;
    const folderResponse = await axios.get(folderUrl, {
      headers: { Authorization: `Bearer ${token}` }
    });

    // Filter for Excel files and find the most recent one
    const excelFiles = folderResponse.data.value
      .filter(item => 
        item.file && 
        item.name.toLowerCase().endsWith('.xlsx') &&
        item.name.includes('_equipment_data')
      )
      .sort((a, b) => new Date(b.lastModifiedDateTime) - new Date(a.lastModifiedDateTime));

    let excelFileName;
    let fileUrl;

    if (excelFiles.length === 0) {
      // If no existing files found, use the standard filename pattern
      const parts = folderName.split('/');
      const subfolderName = parts[parts.length - 1];
      const fileNameBase = subfolderName.replace(/ /g, '_');
      excelFileName = `${fileNameBase}_equipment_data.xlsx`;
      fileUrl = `${GRAPH_ROOT}/users/${ONEDRIVE_USER}/drive/root:/${ONEDRIVE_FOLDER_PATH}/${folderName}/${excelFileName}:/content`;
    } else {
      // Get the most recent Excel file
      const mostRecentFile = excelFiles[0];
      excelFileName = mostRecentFile.name;
      fileUrl = `${GRAPH_ROOT}/users/${ONEDRIVE_USER}/drive/items/${mostRecentFile.id}/content`;
    }

    console.log('Loading Excel file:', excelFileName);

    const response = await axios.get(fileUrl, {
      headers: { Authorization: `Bearer ${token}` },
      responseType: 'arraybuffer'
    });

    const workbook = xlsx.read(response.data, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Get the full range of the sheet to extract header information
    const range = xlsx.utils.decode_range(sheet['!ref']);
    const headerRows = [];
    
    // Extract the first 6 rows as header information
    for (let row = 0; row < 6 && row <= range.e.r; row++) {
      const headerRow = [];
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = xlsx.utils.encode_cell({ r: row, c: col });
        const cell = sheet[cellAddress];
        headerRow.push(cell ? cell.v : '');
      }
      headerRows.push(headerRow);
    }

    // Extract data rows starting from row 7 (index 6)
    const jsonData = xlsx.utils.sheet_to_json(sheet, { 
      range: 6, // Start from row 7 (0-indexed row 6)
      defval: ''
    });

    // Map the data to our expected format manually
    const filteredData = jsonData.map(row => {
      // Get the values from each column
      const values = Object.values(row);
      // Check if the row has a "Selected" column and convert "Yes"/"No" to boolean
      const hasSelectedColumn = values.length > 5;
      const selectedValue = hasSelectedColumn ? (values[5] === 'Yes') : false;
      
      return {
        device_type: values[0] || '',
        manufacturer: values[1] || '',
        model: values[2] || '',
        serial: values[3] || '',
        notes: values[4] || '',
        selected: selectedValue
      };
    }).filter(row => 
      row.device_type || row.manufacturer || row.model || row.serial
    );

    return res.json({ 
      headerRows: headerRows,
      rows: filteredData,
      fileName: excelFileName
    });

  } catch (err) {
    console.error('Detailed Excel fetch error:', {
      message: err.message,
      stack: err.stack,
      response: err.response?.data
    });
    
    res.status(500).json({ 
      error: 'Could not fetch Excel file', 
      details: err.message
    });
  }
});
app.post('/api/save-excel', requireLogin, async (req, res) => {
  try {
    const { folderName, headerRows, rows } = req.body;
    if (!folderName || !rows || !Array.isArray(rows)) {
      return res.status(400).json({ error: 'Missing or invalid folderName or rows' });
    }

    const token = await getAccessToken();
    const parts = folderName.split('/');
    const lastFolderName = parts[parts.length - 1];
    const fileNameBase = lastFolderName.replace(/ /g, '_');
    const now = new Date();
    const timestamp = `${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2,'0')}${now.getDate().toString().padStart(2,'0')}_${now.getHours()}${now.getMinutes()}${now.getSeconds()}`;
    const backupFileName = `${fileNameBase}_equipment_data_${timestamp}.xlsx`;

    // Create the Excel workbook with proper structure
    const wb = xlsx.utils.book_new();
    
    // Create the worksheet data array with proper structure
    const ws_data = [];
    
    // Rows 1-4: Client and Biomedix info
    // Row 1
    ws_data.push(['', '', '', '', '']);
    // Row 2 - Client info in C2, Biomedix in D2
    ws_data.push(['', '', '', '', '']);
    // Row 3 - Client info in C3, Biomedix in D3  
    ws_data.push(['', '', '', '', '']);
    // Row 4 - Client info in C4, Biomedix in D4
    ws_data.push(['', '', '', '', '']);
    // Row 5 - Client info in C5, Biomedix in D5
    ws_data.push(['', '', '', '', '']);
    
    // Add client information from headerRows if available
    if (headerRows && Array.isArray(headerRows) && headerRows.length >= 4) {
      // Client info in column C (rows 2-5)
      ws_data[1][2] = headerRows[0][2] || ''; // C2
      ws_data[2][2] = headerRows[1][2] || ''; // C3
      ws_data[3][2] = headerRows[2][2] || ''; // C4
      ws_data[4][2] = headerRows[3][2] || ''; // C5
    }
    
    // Add BioMedix information in column D (rows 2-5)
    ws_data[1][3] = 'BioMedix Engineering Inc';     // D2
    ws_data[2][3] = '2030 Bristol Circle, Suite 210'; // D3
    ws_data[3][3] = 'Oakville, ON L6H 6P5';         // D4
    ws_data[4][3] = 'Ph.: 416-875-1407';            // D5
    
    // Empty row (row 6)
    ws_data.push(['', '', '', '', '']);
    
    // Row 7: Column headers
    ws_data.push(['Device Type', 'Manufacturer', 'Model', 'Serial Number', 'Notes', 'Selected']);
    
    // Add data rows starting from row 8
    rows.forEach(row => {
      ws_data.push([
        row.device_type || '',
        row.manufacturer || '',
        row.model || '',
        row.serial || '',
        row.notes || '',
        row.selected ? 'Yes' : 'No'
      ]);
    });
    
    const ws = xlsx.utils.aoa_to_sheet(ws_data);
    xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
    const wbBuffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });

    // Save the backup version with timestamp
    const backupUrl = `${GRAPH_ROOT}/users/${ONEDRIVE_USER}/drive/root:/${ONEDRIVE_FOLDER_PATH}/${folderName}/${backupFileName}:/content`;
    await axios.put(backupUrl, wbBuffer, {
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }
    });

    res.json({ 
      success: true,
      message: 'Backup file saved successfully',
      fileName: backupFileName,
      originalFilePreserved: true
    });
  } catch (err) {
    console.error('Error saving backup Excel:', {
      message: err.message,
      response: err.response?.data,
      stack: err.stack
    });
    res.status(500).json({ 
      error: 'Failed to save backup Excel file', 
      details: err.message,
      originalFilePreserved: true
    });
  }
});

async function getAccessToken() {
  const response = await axios.post(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: CLIENT_ID,
      client_secret: CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default',
    }).toString(),
    {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    }
  );
  return response.data.access_token;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Server running on http://localhost:${PORT}`);
});
