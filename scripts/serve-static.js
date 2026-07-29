'use strict';

const http = require('node:http');
const path = require('node:path');
const { readFile, stat } = require('node:fs/promises');

const root = path.resolve(__dirname, '..');
const port = Number(process.env.PORT || 4180);
const mimeTypes = {
    '.css': 'text/css; charset=utf-8',
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png'
};

http.createServer(async (request, response) => {
    try {
        const url = new URL(request.url, `http://127.0.0.1:${port}`);
        const pathname = decodeURIComponent(url.pathname === '/' ? '/index.html' : url.pathname);
        let target = path.resolve(root, `.${pathname}`);
        if (target !== root && !target.startsWith(`${root}${path.sep}`)) {
            response.writeHead(403).end('Forbidden');
            return;
        }
        if ((await stat(target)).isDirectory()) target = path.join(target, 'index.html');
        const content = await readFile(target);
        response.writeHead(200, {
            'Content-Type': mimeTypes[path.extname(target).toLowerCase()] || 'application/octet-stream',
            'Cache-Control': 'no-store'
        });
        response.end(content);
    } catch (error) {
        response.writeHead(error.code === 'ENOENT' ? 404 : 500).end('Not found');
    }
}).listen(port, '127.0.0.1', () => {
    console.log(`BTCwebapp static server: http://127.0.0.1:${port}/`);
});
