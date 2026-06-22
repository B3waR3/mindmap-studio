const express = require('express');
const path = require('path');
const https = require('https');
const http = require('http');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Required for Office Add-ins running in WebView2 (new OneNote for Windows).
// Without these headers, the browser blocks localhost connections from the add-in frame.
app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Private-Network', 'true');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', '*');
  if (req.method === 'OPTIONS') { res.sendStatus(200); return; }
  next();
});

app.use('/taskpane', express.static(path.join(__dirname, 'src', 'taskpane')));
app.use('/mindmap', express.static(path.join(__dirname, 'src', 'mindmap')));
app.use('/assets', express.static(path.join(__dirname, 'assets')));
app.use(express.static(path.join(__dirname, 'src')));

// Try to start HTTPS (required for Office Add-ins in production)
// Run `npm run install-certs` once to generate dev certs
const homePath = process.env.HOME || process.env.USERPROFILE || '';
const certDir = path.join(homePath, '.office-addin-dev-certs');
const keyPath = path.join(certDir, 'localhost.key');
const certPath = path.join(certDir, 'localhost.crt');

if (fs.existsSync(keyPath) && fs.existsSync(certPath)) {
  const options = {
    key: fs.readFileSync(keyPath),
    cert: fs.readFileSync(certPath)
  };
  https.createServer(options, app).listen(PORT, () => {
    console.log(`✅ HTTPS server running at https://localhost:${PORT}`);
    console.log(`   Open manifest.xml in Office to sideload the add-in.`);
  });
} else {
  http.createServer(app).listen(PORT, () => {
    console.log(`⚠️  HTTP server running at http://localhost:${PORT}`);
    console.log(`   For Office Add-ins, HTTPS is required.`);
    console.log(`   Run: npm run install-certs`);
  });
}
