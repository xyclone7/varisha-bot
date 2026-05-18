const express = require('express');
const cors = require('cors');
const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json());

// ── Config ────────────────────────────────────────
const TELEGRAM_TOKEN = process.env.TELEGRAM_TOKEN;
const TELEGRAM_CHAT_ID = process.env.TELEGRAM_CHAT_ID;
const PORT = process.env.PORT || 3000;
const EXCEL_PATH = path.join(__dirname, '..', 'orders.xlsx');

const bot = new TelegramBot(TELEGRAM_TOKEN, { polling: false });

// ── Excel Setup ───────────────────────────────────
const HEADERS = [
  { header: 'No', key: 'no', width: 6 },
  { header: 'Tgl Pesanan', key: 'tanggal', width: 20 },
  { header: 'Order ID', key: 'orderId', width: 16 },
  { header: 'Nama', key: 'nama', width: 22 },
  { header: 'No. WhatsApp', key: 'whatsapp', width: 18 },
  { header: 'Alamat', key: 'alamat', width: 40 },
  { header: 'Kota Tujuan', key: 'kota', width: 18 },
  { header: 'Paket', key: 'paket', width: 10 },
  { header: 'Warna', key: 'warna', width: 30 },
  { header: 'Layanan JNE', key: 'jne', width: 14 },
  { header: 'Harga Barang', key: 'hargaBarang', width: 16 },
  { header: 'Ongkir', key: 'ongkir', width: 14 },
  { header: 'Total', key: 'total', width: 16 },
  { header: 'Status', key: 'status', width: 14 },
];

async function getOrCreateWorkbook() {
  const workbook = new ExcelJS.Workbook();

  if (fs.existsSync(EXCEL_PATH)) {
    await workbook.xlsx.readFile(EXCEL_PATH);
    return workbook;
  }

  // Buat file baru dengan header & styling
  const sheet = workbook.addWorksheet('Pesanan');
  sheet.columns = HEADERS;

  // Style header row
  const headerRow = sheet.getRow(1);
  headerRow.eachCell(cell => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC8826E' } };
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FFA06050' } },
      bottom: { style: 'thin', color: { argb: 'FFA06050' } },
      left: { style: 'thin', color: { argb: 'FFA06050' } },
      right: { style: 'thin', color: { argb: 'FFA06050' } },
    };
  });
  headerRow.height = 32;

  await workbook.xlsx.writeFile(EXCEL_PATH);
  return workbook;
}

async function appendOrder(orderData) {
  const workbook = await getOrCreateWorkbook();
  const sheet = workbook.getWorksheet('Pesanan');

  const lastRow = sheet.lastRow?.number || 1;
  const rowNumber = lastRow + 1;
  const no = lastRow; // row 1 = header, jadi no urut = lastRow

  const newRow = sheet.addRow({
    no,
    tanggal: new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' }),
    orderId: orderData.orderId,
    nama: orderData.nama,
    whatsapp: orderData.whatsapp,
    alamat: orderData.alamat,
    kota: orderData.kota,
    paket: `${orderData.paketQty} pcs`,
    warna: orderData.warna.join(', '),
    jne: orderData.jneService,
    hargaBarang: orderData.hargaBarang,
    ongkir: orderData.ongkir,
    total: orderData.total,
    status: 'Menunggu Pembayaran',
  });

  // Style data rows — zebra striping
  const isEven = (rowNumber % 2 === 0);
  newRow.eachCell(cell => {
    cell.fill = {
      type: 'pattern', pattern: 'solid',
      fgColor: { argb: isEven ? 'FFFFF8F6' : 'FFFFFFFF' },
    };
    cell.border = {
      top: { style: 'hair', color: { argb: 'FFE8DDD6' } },
      bottom: { style: 'hair', color: { argb: 'FFE8DDD6' } },
      left: { style: 'hair', color: { argb: 'FFE8DDD6' } },
      right: { style: 'hair', color: { argb: 'FFE8DDD6' } },
    };
    cell.alignment = { vertical: 'middle', wrapText: true };
  });

  // Format currency cells
  ['hargaBarang', 'ongkir', 'total'].forEach(key => {
    const col = HEADERS.findIndex(h => h.key === key) + 1;
    const cell = newRow.getCell(col);
    cell.numFmt = '"Rp "#,##0';
  });

  // Status cell color
  const statusCol = HEADERS.findIndex(h => h.key === 'status') + 1;
  newRow.getCell(statusCol).fill = {
    type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF3CD' },
  };
  newRow.getCell(statusCol).font = { color: { argb: 'FF856404' }, bold: true };

  newRow.height = 28;

  await workbook.xlsx.writeFile(EXCEL_PATH);
  return no;
}

// ── Telegram Notif ────────────────────────────────
async function sendTelegramNotif(orderData, rowNo) {
  const msg = `
🛍️ *PESANAN BARU — Varisha Hijab*
━━━━━━━━━━━━━━━━━━━━
📋 *Order ID:* \`${orderData.orderId}\`
👤 *Nama:* ${orderData.nama}
📱 *WhatsApp:* [${orderData.whatsapp}](https://wa.me/${orderData.whatsapp.replace(/\D/g,'')}?text=Halo+${encodeURIComponent(orderData.nama)}%2C+pesanan+kamu+sudah+kami+terima!)
📍 *Kota:* ${orderData.kota}
━━━━━━━━━━━━━━━━━━━━
🧕 *Paket:* ${orderData.paketQty} pcs
🎨 *Warna:* ${orderData.warna.join(', ')}
🚚 *JNE:* ${orderData.jneService}
━━━━━━━━━━━━━━━━━━━━
💰 Barang: Rp ${orderData.hargaBarang.toLocaleString('id-ID')}
📦 Ongkir: Rp ${orderData.ongkir.toLocaleString('id-ID')}
✅ *Total: Rp ${orderData.total.toLocaleString('id-ID')}*
━━━━━━━━━━━━━━━━━━━━
📊 Tersimpan di Excel baris #${rowNo}
  `.trim();

  await bot.sendMessage(TELEGRAM_CHAT_ID, msg, {
    parse_mode: 'Markdown',
    reply_markup: {
      inline_keyboard: [[
        { text: '💬 Chat Pembeli', url: `https://wa.me/${orderData.whatsapp.replace(/\D/g,'')}?text=Halo+${encodeURIComponent(orderData.nama)}%2C+terima+kasih+sudah+pesan+Varisha+Hijab!+Berikut+info+pembayarannya+ya+kak+%F0%9F%98%8A` },
        { text: '📥 Download Excel', callback_data: 'download_excel' },
      ]],
    },
  });
}

// ── Bot Commands ──────────────────────────────────
bot.on('polling_error', () => {}); // silent polling errors (kita pakai webhook mode nanti)

// Handle callback download excel
bot.on('callback_query', async (query) => {
  if (query.data === 'download_excel') {
    if (!fs.existsSync(EXCEL_PATH)) {
      return bot.answerCallbackQuery(query.id, { text: 'Belum ada pesanan.' });
    }
    await bot.sendDocument(query.message.chat.id, EXCEL_PATH, {
      caption: `📊 Data pesanan Varisha Hijab — ${new Date().toLocaleDateString('id-ID')}`,
    });
    bot.answerCallbackQuery(query.id);
  }
});

// /download command
bot.onText(/\/download/, async (msg) => {
  if (String(msg.chat.id) !== String(TELEGRAM_CHAT_ID)) return;
  if (!fs.existsSync(EXCEL_PATH)) {
    return bot.sendMessage(msg.chat.id, '📭 Belum ada pesanan masuk.');
  }
  await bot.sendDocument(msg.chat.id, EXCEL_PATH, {
    caption: `📊 Data pesanan Varisha Hijab — ${new Date().toLocaleDateString('id-ID')}`,
  });
});

// /status command
bot.onText(/\/status/, async (msg) => {
  if (String(msg.chat.id) !== String(TELEGRAM_CHAT_ID)) return;
  if (!fs.existsSync(EXCEL_PATH)) {
    return bot.sendMessage(msg.chat.id, '📭 Belum ada pesanan masuk.');
  }
  const wb = await getOrCreateWorkbook();
  const sheet = wb.getWorksheet('Pesanan');
  const total = sheet.lastRow?.number - 1 || 0;
  bot.sendMessage(msg.chat.id, `📊 Total pesanan masuk: *${total}* order\nKetik /download untuk ambil file Excel.`, { parse_mode: 'Markdown' });
});

// /start command
bot.onText(/\/start/, (msg) => {
  if (String(msg.chat.id) !== String(TELEGRAM_CHAT_ID)) return;
  bot.sendMessage(msg.chat.id,
    `👋 *Bot Varisha Hijab aktif!*\n\nCommand tersedia:\n/download — Download file Excel pesanan\n/status — Lihat jumlah pesanan\n\nBot akan otomatis notif setiap ada pesanan baru. 🛍️`,
    { parse_mode: 'Markdown' }
  );
});

// Start polling
if (TELEGRAM_TOKEN) {
  bot.startPolling();
  console.log('🤖 Telegram bot polling aktif');
}

// ── API Endpoint ──────────────────────────────────
app.post('/api/order', async (req, res) => {
  try {
    const data = req.body;

    // Validasi
    if (!data.nama || !data.whatsapp || !data.orderId) {
      return res.status(400).json({ error: 'Data tidak lengkap' });
    }

    // Simpan ke Excel
    const rowNo = await appendOrder(data);

    // Kirim notif Telegram
    await sendTelegramNotif(data, rowNo);

    res.json({ success: true, orderId: data.orderId });
  } catch (err) {
    console.error('Error:', err);
    res.status(500).json({ error: 'Server error' });
  }
});

// Health check
app.get('/', (req, res) => res.json({ status: 'Varisha Bot online ✅' }));

app.listen(PORT, () => console.log(`🚀 Server jalan di port ${PORT}`));
