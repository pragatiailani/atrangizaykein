const express = require("express");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;
const ORDERS_FILE = path.join(__dirname, "orders-log.xlsx");

app.use(express.json());
app.use(express.static(__dirname));

function getRequestIp(req) {
  const fwd = req.headers["x-forwarded-for"];
  if (typeof fwd === "string" && fwd.length) {
    return fwd.split(",")[0].trim();
  }
  return req.socket?.remoteAddress || "Unknown";
}

function appendOrder(order, req) {
  let workbook;
  let worksheet;
  if (fs.existsSync(ORDERS_FILE)) {
    workbook = XLSX.readFile(ORDERS_FILE);
    const sheetName = workbook.SheetNames[0] || "Orders";
    worksheet = workbook.Sheets[sheetName];
  } else {
    workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([
      ["Order ID", "Name", "Items", "Total", "Date Time", "IP"],
    ]);
    XLSX.utils.book_append_sheet(workbook, sheet, "Orders");
    worksheet = sheet;
  }

  const sheetName = workbook.SheetNames[0];
  worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

  const lastDataRow = rows.length > 1 ? rows[rows.length - 1] : null;
  const lastId = lastDataRow ? parseInt(lastDataRow[0], 10) : 0;
  const orderId = Number.isFinite(lastId) ? lastId + 1 : rows.length;

  const orderedAtIso = order.orderedAt || new Date().toISOString();
  const userIp = order.clientIp || getRequestIp(req);
  const itemsSummary = (order.items || [])
    .map((item) => {
      const perItem = item.qty ? Math.round(item.price / item.qty) : item.price;
      const qty = item.qty || 0;
      const price = item.price || 0;
      return `${item.name || ""} x${qty} @ Rs${perItem} (Rs${price})`;
    })
    .join("; ");
  const formattedDate = (() => {
    const d = new Date(orderedAtIso);
    const pad = (n) => String(n).padStart(2, "0");
    const dd = pad(d.getDate());
    const mm = pad(d.getMonth() + 1);
    const yy = String(d.getFullYear()).slice(-2);
    const hh = pad(d.getHours());
    const min = pad(d.getMinutes());
    return `${dd}-${mm}-${yy}, ${hh}:${min}`;
  })();

  const newRow = [
    orderId,
    order.name || "Guest",
    itemsSummary,
    order.totalPrice || 0,
    formattedDate,
    userIp,
  ];

  XLSX.utils.sheet_add_aoa(worksheet, [newRow], { origin: rows.length });
  XLSX.writeFile(workbook, ORDERS_FILE);
  return { orderId, orderedAt: orderedAtIso, userIp };
}

app.post("/api/orders", (req, res) => {
  try {
    const payload = req.body || {};
    if (!Array.isArray(payload.items) || payload.items.length === 0) {
      return res.status(400).json({ error: "No items provided" });
    }
    const result = appendOrder(payload, req);
    res.json({ ok: true, ...result });
  } catch (error) {
    console.error("Failed to append order", error);
    res.status(500).json({ error: "Failed to record order" });
  }
});

app.listen(PORT, () => {
  console.log(`FoodFest server running on http://localhost:${PORT}`);
});
