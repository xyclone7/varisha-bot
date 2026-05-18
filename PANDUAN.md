# 🚀 Panduan Setup Varisha Bot

## Struktur Folder
```
varisha-bot/
├── server/
│   └── index.js        ← Node.js server + Telegram bot
├── public/
│   └── index.html      ← Landing page (deploy ke Netlify/Vercel)
├── orders.xlsx         ← Auto-dibuat saat pesanan pertama masuk
├── package.json
├── .env.example
└── .gitignore
```

---

## STEP 1 — Buat Telegram Bot

1. Buka Telegram → cari **@BotFather**
2. Ketik `/newbot`
3. Isi nama bot: `Varisha Hijab Bot`
4. Isi username: `varisha_hijab_bot` (harus unik, tambahin angka kalau sudah dipakai)
5. Copy **TOKEN** yang dikasih BotFather → simpan

**Dapatkan Chat ID kamu:**
1. Cari bot **@userinfobot** di Telegram
2. Start bot tersebut → dia akan kirim Chat ID kamu
3. Copy angkanya → simpan

---

## STEP 2 — Deploy Server ke Railway

1. Buka **railway.app** → Login dengan GitHub
2. Klik **New Project** → **Deploy from GitHub repo**
3. Upload / push folder `varisha-bot` ke GitHub dulu:
   ```bash
   git init
   git add .
   git commit -m "init varisha bot"
   git remote add origin https://github.com/USERNAME/varisha-bot.git
   git push -u origin main
   ```
4. Di Railway → pilih repo `varisha-bot`
5. Railway akan auto-detect Node.js dan deploy

**Set Environment Variables di Railway:**
- Klik project → **Variables** → tambahkan:
  ```
  TELEGRAM_TOKEN   = token dari BotFather
  TELEGRAM_CHAT_ID = chat id kamu
  PORT             = 3000
  ```

6. Setelah deploy → Railway kasih URL seperti:
   `https://varisha-bot-production.up.railway.app`

---

## STEP 3 — Update URL di HTML

Buka `public/index.html`, cari baris ini di bagian script:
```js
const SERVER_URL = 'https://GANTI-URL-RAILWAY-KAMU.railway.app';
```
Ganti dengan URL Railway kamu yang asli.

---

## STEP 4 — Deploy HTML ke Netlify

1. Buka **netlify.com** → Login
2. Drag & drop folder `public/` ke dashboard Netlify
3. Netlify kasih URL gratis, contoh: `https://varisha-hijab.netlify.app`
4. Selesai! Landing page kamu live 🎉

---

## STEP 5 — Test

1. Buka landing page kamu
2. Isi form → klik "Lanjut ke Pembayaran"
3. Cek Telegram kamu → harusnya langsung dapat notif
4. Di notif ada tombol **💬 Chat Pembeli** → tinggal klik untuk balas via WA
5. Ketik `/download` di Telegram → bot kirim file Excel

---

## Telegram Bot Commands

| Command | Fungsi |
|---------|--------|
| `/start` | Cek bot aktif |
| `/download` | Download file Excel semua pesanan |
| `/status` | Lihat total jumlah pesanan |

---

## Flow Lengkap Setelah Live

```
Pembeli isi form → klik Pesan
        ↓
Server Railway terima data
        ↓
    ┌───┴───┐
    ↓       ↓
Tulis ke   Kirim notif
Excel      Telegram ke kamu
           (ada tombol Chat Pembeli via WA)
        ↓
Kamu klik tombol → WA ke pembeli
→ Kirim info pembayaran manual
```

---

## Download Excel

Kapan saja ketik `/download` di Telegram → bot langsung kirim file `orders.xlsx`.

File Excel format:
| No | Tgl Pesanan | Order ID | Nama | WhatsApp | Alamat | Kota | Paket | Warna | JNE | Harga | Ongkir | Total | Status |

---

## Troubleshooting

**Bot tidak kirim notif?**
- Cek TELEGRAM_TOKEN dan TELEGRAM_CHAT_ID di Railway Variables sudah benar
- Pastikan kamu sudah `/start` bot di Telegram

**CORS error di browser?**
- Pastikan URL di `SERVER_URL` di HTML sudah benar (tanpa trailing slash)
- Railway harus sudah running (bukan sleep)

**Railway sleep / lambat?**
- Free tier Railway bisa sleep setelah inaktif
- Upgrade ke $5/bulan untuk always-on, atau pakai **Render.com** (free tier lebih stabil untuk hobby project)
