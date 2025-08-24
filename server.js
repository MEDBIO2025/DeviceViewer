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
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

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

    const parts = folderName.split('/');
    const fileNameBase = parts.join('_');
    const excelFileName = `${fileNameBase}_equipment_data.xlsx`;

    const fileUrl = `${GRAPH_ROOT}/users/${ONEDRIVE_USER}/drive/root:/${ONEDRIVE_FOLDER_PATH}/${folderName}/${excelFileName}:/content`;

    const response = await axios.get(fileUrl, {
      headers: { Authorization: `Bearer ${token}` },
      responseType: 'arraybuffer'
    });

    const workbook = xlsx.read(response.data, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const sheetData = xlsx.utils.sheet_to_json(sheet, { 
      range: 6,      // skip first 6 rows
      defval: '',    // default empty string for empty cells
    });

    return res.json({ 
      rows: sheetData,
      fileName: excelFileName
    });

  } catch (err) {
    console.error('Error fetching Excel:', err.response?.data || err.message);
    res.status(500).json({ error: 'Could not fetch Excel file', details: err.message });
  }
});


app.post('/api/save-excel', requireLogin, async (req, res) => {
  try {
    const { folderName, rows } = req.body;
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

    // Create the Excel workbook
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.json_to_sheet(rows);
    xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
    const wbBuffer = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });

    // Save ONLY the backup version with timestamp
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
