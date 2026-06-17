const https = require("https");
const fs = require("fs");
const path = require("path");

const certDir = path.join(require("os").homedir(), ".office-addin-dev-certs");
const options = {
  key: fs.readFileSync(path.join(certDir, "localhost.key")),
  cert: fs.readFileSync(path.join(certDir, "localhost.crt")),
};

const mimeTypes = {
  ".html": "text/html",
  ".js": "application/javascript",
  ".css": "text/css",
  ".json": "application/json",
  ".png": "image/png",
  ".xml": "application/xml",
};

const server = https.createServer(options, (req, res) => {
  let filePath = path.join(__dirname, req.url === "/" ? "/src/taskpane.html" : req.url);
  const ext = path.extname(filePath);
  const contentType = mimeTypes[ext] || "application/octet-stream";

  fs.readFile(filePath, (err, content) => {
    if (err) {
      res.writeHead(404);
      res.end("Not found");
    } else {
      res.writeHead(200, { "Content-Type": contentType });
      res.end(content);
    }
  });
});

server.listen(3000, () => {
  console.log("Serving add-in at https://localhost:3000");
  console.log("Press Ctrl+C to stop.");
});
