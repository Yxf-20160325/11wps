const express = require('express');
const path = require('path');
const app = express();
const PORT = 3000;

// 访问日志中间件
app.use((req, res, next) => {
  const timestamp = new Date().toLocaleString('zh-CN');
  const ip = req.ip || req.connection.remoteAddress;
  const method = req.method;
  const url = req.originalUrl;
  console.log(`[${timestamp}] ${ip} - ${method} ${url} - 有人进入网站`);
  next();
});

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`WPS Office Web App running at http://localhost:${PORT}`);
});
