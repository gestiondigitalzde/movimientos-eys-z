const http = require("http");
const fs = require("fs");
const path = require("path");
const { exec } = require("child_process");

const PORT = 5500;
const DIR = __dirname;

const server = http.createServer((req, res) => {
  const filePath = path.join(DIR, "index.html");
  fs.readFile(filePath, (err, content) => {
    if (err) {
      res.writeHead(500);
      res.end("Error leyendo index.html");
      return;
    }
    res.writeHead(200, {
      "Content-Type": "text/html; charset=utf-8",
      "Cache-Control": "no-store, no-cache, must-revalidate, proxy-revalidate"
    });
    res.end(content);
  });
});

server.on("error", (e) => {
  if (e.code === "EADDRINUSE") {
    exec(`start http://localhost:${PORT}`);
    process.exit(0);
  }
});

server.listen(PORT, () => {
  console.log(`Servidor de pruebas activo en: http://localhost:${PORT}`);
  exec(`start http://localhost:${PORT}`);
});

