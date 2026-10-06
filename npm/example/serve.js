// Serves the npm package folder so the example page can load ../dist.
// Usage: npm run example  (then open the printed URL)
const http = require('http');
const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const port = Number(process.env.PORT) || 8080;
const types = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
};

http
  .createServer((req, res) => {
    const urlPath = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
    if (urlPath === '/') {
      res.writeHead(302, { Location: '/example/' });
      return res.end();
    }
    let file = path.join(root, urlPath);
    if (file !== root && !file.startsWith(root + path.sep)) {
      res.writeHead(403);
      return res.end();
    }
    if (urlPath.endsWith('/')) file = path.join(file, 'index.html');
    fs.readFile(file, (err, data) => {
      if (err) {
        res.writeHead(404);
        return res.end('Not found');
      }
      res.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream' });
      res.end(data);
    });
  })
  .listen(port, () => console.log(`Example running at http://localhost:${port}/example/`));
