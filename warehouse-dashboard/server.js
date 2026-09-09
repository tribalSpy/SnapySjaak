require('dotenv').config();
const express = require('express');
const { createWarehouseDashboardRouter } = require('./src/router');

const app = express();
const PORT = process.env.PORT || 4500;

app.use('/', createWarehouseDashboardRouter());

app.listen(PORT, () => {
  console.log(`Warehouse dashboard running on http://localhost:${PORT}`);
  console.log(`Reading from backend: ${process.env.WAREHOUSE_BACKEND_URL || 'http://127.0.0.1:4000'}`);
});
