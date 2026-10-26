const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const BASE_DIR = __dirname;
const UPLOADS_DIR = path.join(BASE_DIR, 'uploads');

// Ensure storage directory exists
if (!fs.existsSync(UPLOADS_DIR)) {
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
}

const MIME_TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.js': 'application/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.xls': 'application/vnd.ms-excel',
  '.csv': 'text/csv; charset=utf-8',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon'
};

function sendJson(res, statusCode, data) {
  res.writeHead(statusCode, {
    'Content-Type': 'application/json; charset=utf-8',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  });
  res.end(JSON.stringify(data));
}

const server = http.createServer((req, res) => {
  const parsedUrl = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  let pathname = decodeURIComponent(parsedUrl.pathname);

  // Handle CORS Preflight
  if (req.method === 'OPTIONS') {
    res.writeHead(204, {
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'GET, POST, DELETE, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
    return res.end();
  }

  // --- API: List Stored Files ---
  if (req.method === 'GET' && pathname === '/api/files') {
    fs.readdir(UPLOADS_DIR, { withFileTypes: true }, (err, entries) => {
      if (err) {
        return sendJson(res, 500, { error: 'Failed to read storage directory: ' + err.message });
      }

      const files = entries
        .filter(e => e.isFile() && !e.name.startsWith('.'))
        .map(e => {
          const filePath = path.join(UPLOADS_DIR, e.name);
          const stats = fs.statSync(filePath);
          return {
            name: e.name,
            size: stats.size,
            mtime: stats.mtime,
            url: `/uploads/${encodeURIComponent(e.name)}`
          };
        })
        .sort((a, b) => new Date(b.mtime) - new Date(a.mtime));

      return sendJson(res, 200, { files, storagePath: UPLOADS_DIR });
    });
    return;
  }

  // --- API: Upload / Save File to Storage ---
  if (req.method === 'POST' && pathname === '/api/upload') {
    let body = '';
    req.on('data', chunk => {
      body += chunk;
      // Max 100MB payload guard
      if (body.length > 100 * 1024 * 1024) {
        sendJson(res, 413, { error: 'Payload too large (limit is 100MB)' });
        req.destroy();
      }
    });

    req.on('end', () => {
      try {
        const payload = JSON.parse(body);
        const items = Array.isArray(payload) ? payload : [payload];
        const saved = [];

        for (const item of items) {
          if (!item.filename || !item.contentBase64) continue;

          // Sanitize filename to avoid path traversal
          const safeName = path.basename(item.filename).replace(/[^a-zA-Z0-9._\-]/g, '_');
          const targetPath = path.join(UPLOADS_DIR, safeName);
          const buffer = Buffer.from(item.contentBase64, 'base64');

          fs.writeFileSync(targetPath, buffer);
          saved.push({
            name: safeName,
            size: buffer.length,
            url: `/uploads/${encodeURIComponent(safeName)}`
          });
        }

        console.log(`[Storage] Saved ${saved.length} file(s) into ${UPLOADS_DIR}`);
        return sendJson(res, 200, { success: true, saved });
      } catch (err) {
        console.error('[Storage Upload Error]', err);
        return sendJson(res, 400, { error: 'Invalid upload payload: ' + err.message });
      }
    });
    return;
  }

  // --- API: Delete Stored File ---
  if (req.method === 'DELETE' && pathname.startsWith('/api/files/')) {
    const filename = path.basename(decodeURIComponent(pathname.replace('/api/files/', '')));
    const targetPath = path.join(UPLOADS_DIR, filename);

    if (fs.existsSync(targetPath)) {
      try {
        fs.unlinkSync(targetPath);
        console.log(`[Storage] Deleted file: ${filename}`);
        return sendJson(res, 200, { success: true, deleted: filename });
      } catch (err) {
        return sendJson(res, 500, { error: 'Failed to delete file: ' + err.message });
      }
    } else {
      return sendJson(res, 404, { error: 'File not found in storage' });
    }
  }

  // --- Static File Serving ---
  if (pathname === '/' || pathname === '') {
    pathname = '/index.html';
  }

  const filePath = path.normalize(path.join(BASE_DIR, pathname));

  // Security check: ensure path is within BASE_DIR
  if (!filePath.startsWith(BASE_DIR)) {
    res.writeHead(403, { 'Content-Type': 'text/plain' });
    res.end('403 Forbidden');
    return;
  }

  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      res.end('404 Not Found');
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_TYPES[ext] || 'application/octet-stream';

    const headers = {
      'Content-Type': contentType,
      'Content-Length': stats.size,
      'Access-Control-Allow-Origin': '*'
    };

    if (ext === '.xlsx' || ext === '.xls' || ext === '.csv') {
      headers['Content-Disposition'] = `attachment; filename="${path.basename(filePath)}"`;
    }

    res.writeHead(200, headers);
    fs.createReadStream(filePath).pipe(res);
  });
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Student Attendance Intelligence Dashboard running at http://localhost:${PORT}/`);
  console.log(`Server Storage Directory: ${UPLOADS_DIR}`);
});
