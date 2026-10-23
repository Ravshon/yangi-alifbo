/**
 * Yangi Alifbo - REST API Server
 * 
 * Muallif: Ravshon (https://github.com/Ravshon)
 * Litsenziya: MIT (GitHub profilini ko'rsatish shart)
 * 
 * Ishga tushirish: node api_server.js
 * Default port: 3000
 */

const http = require('http');
const url = require('url');
const YangiAlifbo = require('./yangi-alifbo.js');

const PORT = process.env.PORT || 3000;

function setCorsHeaders(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

const server = http.createServer((req, res) => {
  setCorsHeaders(res);

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const parsedUrl = url.parse(req.url, true);
  const pathname = parsedUrl.pathname;

  // 1. Health check & Info
  if (pathname === '/' || pathname === '/api' || pathname === '/api/health') {
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify({
      status: 'online',
      project: "O'zbekcha Yangi Alifbo API",
      author: 'Ravshon',
      github: 'https://github.com/Ravshon',
      license: 'MIT (Attribution Required)',
      endpoints: {
        convert_get: '/api/convert?text=...&mode=cyrToNewLat',
        convert_post: '/api/convert (JSON body: {"text": "...", "mode": "cyrToNewLat"})'
      },
      modes: [
        'cyrToNewLat (Kirill -> Yangi Lotin: Ö, Ğ, Ş, Ç)',
        'cyrToOldLat (Kirill -> Eski Lotin: O\', G\', Sh, Ch)',
        'latToCyr (Lotin -> Kirill)',
        'oldToNewLat (Eski Lotin -> Yangi Lotin)',
        'newToOldLat (Yangi Lotin -> Eski Lotin)'
      ]
    }, null, 2));
    return;
  }

  // 2. Conversion endpoint
  if (pathname === '/api/convert') {
    if (req.method === 'GET') {
      const text = parsedUrl.query.text || '';
      const mode = parsedUrl.query.mode || 'cyrToNewLat';

      const result = YangiAlifbo.convert(text, mode);
      res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
      res.end(JSON.stringify({
        success: true,
        mode: mode,
        input: text,
        result: result.result,
        stats: {
          chars: text.length,
          replacements: result.count
        },
        author: 'https://github.com/Ravshon'
      }));
      return;
    }

    if (req.method === 'POST') {
      let body = '';
      req.on('data', chunk => {
        body += chunk.toString();
        // Limit max body size to 5MB
        if (body.length > 5 * 1024 * 1024) {
          res.writeHead(413, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Matn hajmi juda katta (maksimal 5MB)' }));
          req.connection.destroy();
        }
      });

      req.on('end', () => {
        try {
          const data = JSON.parse(body || '{}');
          const text = data.text || '';
          const mode = data.mode || 'cyrToNewLat';

          const result = YangiAlifbo.convert(text, mode);
          res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
          res.end(JSON.stringify({
            success: true,
            mode: mode,
            result: result.result,
            stats: {
              chars: text.length,
              replacements: result.count
            },
            author: 'https://github.com/Ravshon'
          }));
        } catch (e) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: "Noto'g'ri JSON formati: " + e.message }));
        }
      });
      return;
    }
  }

  // 404
  res.writeHead(404, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify({ error: 'Endpoint topilmadi. /api/convert ga murojaat qiling.' }));
});

server.listen(PORT, () => {
  console.log(`🚀 Yangi Alifbo API serveri ishga tushdi: http://localhost:${PORT}`);
  console.log(`📖 Hujjatlar: http://localhost:${PORT}/api`);
});
