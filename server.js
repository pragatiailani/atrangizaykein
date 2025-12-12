const express = require("express");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;
const ORDERS_FILE = path.join(__dirname, "orders-log.xlsx");
const MENU_FILE = path.join(__dirname, "menu-items.xlsx");

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
  const header = ["Order ID", "Name", "Items", "Total", "Date Time", "IP"];
  let workbook;
  let worksheet;
  if (fs.existsSync(ORDERS_FILE)) {
    workbook = XLSX.readFile(ORDERS_FILE);
    const sheetName = workbook.SheetNames[0] || "Orders";
    worksheet = workbook.Sheets[sheetName];
  } else {
    workbook = XLSX.utils.book_new();
    worksheet = XLSX.utils.aoa_to_sheet([header]);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Orders");
  }

  const sheetName = workbook.SheetNames[0];
  worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils
    .sheet_to_json(worksheet, { header: 1, defval: "" })
    .filter((row) => row.some((cell) => cell !== ""));
  const dataRows = rows.slice(1); // skip header
  const lastId = dataRows.reduce((max, row) => {
    const id = parseInt(row[0], 10);
    return Number.isFinite(id) && id > max ? id : max;
  }, 0);
  const orderId = lastId + 1;

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

function readMenuFile() {
  if (!fs.existsSync(MENU_FILE)) {
    throw new Error("menu-items.xlsx not found");
  }
  const workbook = XLSX.readFile(MENU_FILE);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const metaRow = rows[0] || [];
  const headerRow = rows[1] || [];
  const dataRows = rows.slice(2);
  const headerMap = headerRow.map((h) => String(h || "").trim().toLowerCase());

  const items = dataRows
    .map((row) => {
      const get = (key) => {
        const idx = headerMap.indexOf(key);
        return idx >= 0 ? row[idx] : "";
      };
      const key = String(get("key") || "").trim();
      if (!key) return null;
      return {
        key,
        emoji: get("emoji") || "",
        name: get("name") || "",
        description: get("description") || "",
        maxPrice: Number(get("maxprice")) || 0,
      };
    })
    .filter(Boolean);

  return {
    meta: { stallName: metaRow[0] || "", festName: metaRow[1] || "" },
    items,
  };
}

function writeMenuFile(meta, items) {
  const headerRow = ["Key", "Emoji", "Name", "Description", "MaxPrice"];
  const aoa = [
    [meta?.stallName || "", meta?.festName || ""],
    headerRow,
    ...items.map((item) => [
      item.key || "",
      item.emoji || "",
      item.name || "",
      item.description || "",
      Number(item.maxPrice) || 0,
    ]),
  ];
  const sheet = XLSX.utils.aoa_to_sheet(aoa);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Menu");
  XLSX.writeFile(workbook, MENU_FILE);
}

function readOrdersFile() {
  const header = ["Order ID", "Name", "Items", "Total", "Date Time", "IP"];
  if (!fs.existsSync(ORDERS_FILE)) {
    return { headers: header, rows: [] };
  }
  const workbook = XLSX.readFile(ORDERS_FILE);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const [, ...dataRows] = raw;
  const rows = dataRows
    .filter((row) => row.some((cell) => cell !== ""))
    .map((row) => ({
      orderId: row[0],
      name: row[1],
      items: row[2],
      total: Number(row[3]) || 0,
      dateTime: row[4],
      ip: row[5],
    }));
  return { headers: header, rows };
}

function writeOrdersFile(rows) {
  const header = ["Order ID", "Name", "Items", "Total", "Date Time", "IP"];
  const aoa = [header, ...rows.map((r) => [r.orderId, r.name, r.items, r.total, r.dateTime, r.ip])];
  const sheet = XLSX.utils.aoa_to_sheet(aoa);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Orders");
  XLSX.writeFile(workbook, ORDERS_FILE);
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

app.get("/api/orders", (req, res) => {
  try {
    const orders = readOrdersFile();
    res.json(orders);
  } catch (error) {
    console.error("Failed to read orders", error);
    res.status(500).json({ error: "Failed to read orders" });
  }
});

app.delete("/api/orders/:orderId", (req, res) => {
  try {
    const orderId = String(req.params.orderId || "").trim();
    if (!orderId) return res.status(400).json({ error: "Order ID required" });
    const orders = readOrdersFile();
    const nextRows = orders.rows.filter(
      (row) => String(row.orderId) !== orderId
    );
    if (nextRows.length === orders.rows.length) {
      return res.status(404).json({ error: "Order not found" });
    }
    writeOrdersFile(nextRows);
    res.json({ ok: true, deleted: orderId });
  } catch (error) {
    console.error("Failed to delete order", error);
    res.status(500).json({ error: "Failed to delete order" });
  }
});

app.delete("/api/orders", (req, res) => {
  try {
    writeOrdersFile([]);
    res.json({ ok: true, cleared: true });
  } catch (error) {
    console.error("Failed to clear orders", error);
    res.status(500).json({ error: "Failed to clear orders" });
  }
});

app.get("/api/menu", (req, res) => {
  try {
    const menu = readMenuFile();
    res.json(menu);
  } catch (error) {
    console.error("Failed to read menu", error);
    res.status(500).json({ error: "Failed to read menu" });
  }
});

app.post("/api/menu", (req, res) => {
  try {
    const { key, emoji = "", name = "", description = "", maxPrice = 0 } =
      req.body || {};
    const safeKey = String(key || "").trim();
    if (!safeKey) {
      return res.status(400).json({ error: "Key is required" });
    }
    const menu = readMenuFile();
    if (menu.items.some((item) => item.key === safeKey)) {
      return res.status(409).json({ error: "Key already exists" });
    }
    const newItem = {
      key: safeKey,
      emoji: String(emoji || ""),
      name: String(name || ""),
      description: String(description || ""),
      maxPrice: Number(maxPrice) || 0,
    };
    const nextItems = [...menu.items, newItem];
    writeMenuFile(menu.meta, nextItems);
    res.status(201).json({ ok: true, item: newItem });
  } catch (error) {
    console.error("Failed to add menu item", error);
    res.status(500).json({ error: "Failed to add menu item" });
  }
});

app.put("/api/menu/:key", (req, res) => {
  try {
    const key = String(req.params.key || "").trim();
    if (!key) return res.status(400).json({ error: "Key is required" });
    const menu = readMenuFile();
    const idx = menu.items.findIndex((item) => item.key === key);
    if (idx === -1) {
      return res.status(404).json({ error: "Menu item not found" });
    }
    const payload = req.body || {};
    const updated = {
      ...menu.items[idx],
      emoji: payload.emoji ?? menu.items[idx].emoji,
      name: payload.name ?? menu.items[idx].name,
      description: payload.description ?? menu.items[idx].description,
      maxPrice:
        payload.maxPrice !== undefined
          ? Number(payload.maxPrice) || 0
          : menu.items[idx].maxPrice,
    };
    const nextItems = [...menu.items];
    nextItems[idx] = updated;
    writeMenuFile(menu.meta, nextItems);
    res.json({ ok: true, item: updated });
  } catch (error) {
    console.error("Failed to update menu item", error);
    res.status(500).json({ error: "Failed to update menu item" });
  }
});

app.delete("/api/menu/:key", (req, res) => {
  try {
    const key = String(req.params.key || "").trim();
    if (!key) return res.status(400).json({ error: "Key is required" });
    const menu = readMenuFile();
    const nextItems = menu.items.filter((item) => item.key !== key);
    if (nextItems.length === menu.items.length) {
      return res.status(404).json({ error: "Menu item not found" });
    }
    writeMenuFile(menu.meta, nextItems);
    res.json({ ok: true, deleted: key });
  } catch (error) {
    console.error("Failed to delete menu item", error);
    res.status(500).json({ error: "Failed to delete menu item" });
  }
});

app.listen(PORT, () => {
  console.log(`FoodFest server running on http://localhost:${PORT}`);
});
