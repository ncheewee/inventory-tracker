# 📦 Inventory Tracker

A mobile-first web app for warehouse/facility workers to capture stock movements via camera + OCR, with Telegram notifications and weekly reports.

Built as a **single-file HTML frontend** + **Google Apps Script backend** — no server, no hosting cost, deploys in minutes.

---

## How it works

```
Worker phone                     Apps Script Backend
┌──────────┐    HTTPS POST       ┌─────────────────────┐
│ index.html│ ←────────────────→ │  Code.gs (Web App)   │
│          │                     │                       │
│ 1. PIN   │                     │  • Verify PIN         │
│ 2. Choose│                     │  • Call Gemini Vision │
│    In/Out│                     │  • Upload photo→Drive │
│ 3. Snap  │                     │  • Save to Sheet      │
│    photo │                     │  • Push Telegram      │
│ 4. Verify│                     │  • Update inventory   │
│    & save│                     └─────────────────────┘
│ 5. Done ✓│                           │       │       │
└──────────┘                    ┌──────┘       │       └──────┐
                                ▼              ▼              ▼
                         Google Sheet    Google Drive    Telegram
                         (users, tx,     (photos)        (alerts +
                          inventory)                      weekly report)
```

---

## Features

### Worker flow
- **PIN login** — simple 4-digit keypad, no typing
- **Stock In / Stock Out** — two big icon buttons
- **Camera capture** — uses phone's rear camera
- **OCR via Gemini Vision** — reads item name + quantity from photo automatically
- **Verify screen** — worker confirms/edits detected item + qty before saving
- **Success confirmation** — shows what was recorded + Telegram sent status

### Admin features
- **User management** — add/delete users, set Worker/Admin roles
- **Transaction log** — full history with timestamps
- **Inventory levels** — current stock with low-stock warnings (≤ 20 units)
- **Manual report trigger** — send weekly report on demand

### Notifications
- **Per-transaction Telegram push** — instant alert on every stock in/out
- **Weekly report** — auto-sent every Monday 9am via Apps Script time-driven trigger:
  - Stock movement summary (in/out totals)
  - Current inventory levels
  - Low stock alerts
  - Transaction count by user

---

## File structure

```
inventory-tracker/
├── index.html              # Frontend (single-file app)
├── apps-script/
│   ├── Code.gs             # Backend (paste into Apps Script)
│   └── Setup.md            # Step-by-step deploy guide
└── README.md               # This file
```

---

## Quick start

### Option A: Demo mode (no backend)
1. Open `index.html` in a browser.
2. Works immediately with mock data.
3. Demo PINs: `1234` (Worker) or `9999` (Admin).

### Option B: Full deployment
Follow **[apps-script/Setup.md](apps-script/Setup.md)** for the complete guide.

Summary:
1. Create a Google Sheet with 3 tabs (users, transactions, inventory).
2. Paste `Code.gs` into Apps Script.
3. Get a free Gemini API key + create a Telegram bot.
4. Run `setup()` to store config + create weekly trigger.
5. Deploy as Web App (access: Anyone).
6. Paste the Web App URL into `index.html` → `APPS_SCRIPT_URL`.
7. Done.

---

## Tech stack

| Layer | Technology |
|-------|-----------|
| Frontend | Single-file HTML/CSS/JS (no framework, no build) |
| Backend | Google Apps Script (Web App) |
| Database | Google Sheets (3 tabs) |
| OCR | Google Gemini Vision API (gemini-2.0-flash, free tier) |
| Photo storage | Google Drive (auto-created folder) |
| Notifications | Telegram Bot API |
| Weekly trigger | Apps Script time-driven trigger (Mon 9am) |

---

## Security notes

- **Gemini API key + Telegram bot token** stored in Apps Script `PropertiesService` — never exposed to the client.
- **PIN auth** — simple by design (non-technical workers). For higher security, add Google account OAuth or longer PINs.
- **Admin actions** verified server-side (admin PIN required for user management).
- **Photo sharing** — photos uploaded to Drive with "anyone with link can view" for easy verification. Restrict to specific users if needed.
- **Web app access** set to "Anyone" so workers' phones can reach it without Google login. The backend only responds to valid actions.

---

## Customization

| What | Where |
|------|-------|
| Low stock threshold | `LOW_STOCK_THRESHOLD` in both `index.html` and `Code.gs` |
| Gemini model | `GEMINI_MODEL` in `Code.gs` |
| Drive folder name | `DRIVE_FOLDER_NAME` in `Code.gs` |
| Weekly report day/time | `setup()` or `resetTrigger()` in `Code.gs` |
| App colors | CSS `:root` variables in `index.html` |
