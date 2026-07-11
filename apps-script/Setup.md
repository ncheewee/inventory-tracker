# Setup Guide — Inventory Tracker Backend

## Overview

The backend is a Google Apps Script Web App bound to a Google Sheet. It handles:
- PIN-based login
- Gemini Vision OCR (API key stored server-side)
- Photo upload to Google Drive
- Telegram push notifications
- Weekly report (auto-triggered every Monday 9am)

---

## Step 1: Create the Google Sheet

1. Go to [sheets.new](https://sheets.new) — create a new spreadsheet.
2. Name it **Inventory Tracker**.
3. Create 3 tabs (bottom-left `+`):
   - `users`
   - `transactions`
   - `inventory`

> You don't need to add headers manually — `setup()` will do it. But the tab names must match exactly.

---

## Step 2: Add the Apps Script

1. In the spreadsheet: **Extensions → Apps Script**.
2. Delete any code in `Code.gs`.
3. Open `apps-script/Code.gs` from this project, copy the entire contents.
4. Paste into the Apps Script editor.
5. Click **Save** (💾).

---

## Step 3: Get your API keys

### Gemini API Key (free tier)
1. Go to [Google AI Studio](https://aistudio.google.com/apikey).
2. Sign in with your Google account.
3. Click **Create API Key**.
4. Copy the key (starts with `AIza...`).

### Telegram Bot
1. Open Telegram, search for **@BotFather**.
2. Send `/newbot` → follow prompts to name it (e.g. "Inventory Tracker Bot").
3. Copy the **bot token** (looks like `123456:ABC-DEF...`).

### Telegram Chat ID
1. Send any message (e.g. `/start`) to your new bot in Telegram.
2. Open this URL in a browser (replace `<BOT_TOKEN>`):
   ```
   https://api.telegram.org/bot<BOT_TOKEN>/getUpdates
   ```
3. Look for `"chat":{"id":XXXXXXXXX}` in the JSON response — that number is your chat ID.
   - For a group chat, the ID is negative (e.g. `-100123456789`).

---

## Step 4: Configure & run setup()

1. In the Apps Script editor, find the `setup()` function.
2. Edit the `config` object with your keys:

```javascript
var config = {
  GEMINI_API_KEY:     'AIza...',        // your Gemini key
  TELEGRAM_BOT_TOKEN: '123456:ABC...',  // your bot token
  TELEGRAM_CHAT_ID:   '123456789',      // your chat ID
};
```

3. Select `setup` from the function dropdown at the top.
4. Click **Run**.
5. **Authorize** when prompted (Review permissions → your Google account → Advanced → Go to project → Allow).
6. Check **Execution log** — you should see:
   ```
   ✅ Setup complete. Config stored, trigger created, sheet initialized.
   ```

This will:
- Store your keys securely in `PropertiesService` (not in the code).
- Create sheet headers in all 3 tabs.
- Seed 2 default users: `Admin` (PIN 9999) and `Worker` (PIN 1234).
- Create a weekly Monday 9am trigger for the report.

---

## Step 5: Deploy as Web App

1. Click **Deploy → New deployment** (top-right).
2. Click the gear ⚙️ → **Web app**.
3. Settings:
   - **Description:** Inventory Tracker API
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy**.
5. **Authorize** again if prompted.
6. Copy the **Web app URL** (looks like `https://script.google.com/macros/s/AKfyc.../exec`).

---

## Step 6: Connect the frontend

1. Open `index.html` in this project.
2. Find this line near the top of the `<script>`:

```javascript
const APPS_SCRIPT_URL = '';
```

3. Paste your Web app URL:

```javascript
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfyc.../exec';
```

4. Save. The DEMO MODE banner will disappear — you're live.

---

## Step 7: Test

1. Open `index.html` in a mobile browser (or deploy to GitHub Pages).
2. Log in with PIN `9999` (admin) or `1234` (worker).
3. Try a Stock In → camera → snap → verify → confirm.
4. Check:
   - ✅ Transaction appears in the Google Sheet `transactions` tab.
   - ✅ Inventory updated in `inventory` tab.
   - ✅ Photo saved in Google Drive folder `InventoryTrackerPhotos`.
   - ✅ Telegram message received.
   - ✅ Gemini OCR reads the item name + quantity (test with a real product label).

---

## Adding Users

1. Log in as admin (PIN 9999).
2. Admin Panel → Users → Add User.
3. Enter name, 4-digit PIN, role (Worker/Admin).

Users can also be edited directly in the `users` sheet tab.

---

## Changing the weekly report day/time

In Apps Script, delete the old trigger and create a new one:

```javascript
function resetTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.FRIDAY)  // change day here
    .atHour(17)                             // change hour here (24h)
    .create();
}
```

Run `resetTrigger()` once.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "Invalid PIN" but PIN is correct | Check the `users` tab — ensure `active` column is `TRUE` |
| Camera doesn't open | Use HTTPS (GitHub Pages or `localhost`). Camera requires secure context. |
| OCR returns error | Verify Gemini API key in PropertiesService. Check Execution log in Apps Script. |
| Telegram not sent | Verify bot token + chat ID. Ensure you sent `/start` to the bot first. |
| Photo not saved | Check Drive permissions were granted during authorization. |
| CORS error in browser | Apps Script `doPost` returns JSON via `ContentService` — no CORS issues. If seeing errors, ensure URL ends with `/exec` not `/dev`. |
| Weekly report not arriving | Check Triggers in Apps Script (left sidebar ⏰). Check execution history for errors. |

---

## Files

| File | Purpose |
|------|---------|
| `index.html` | Frontend — single-file web app |
| `apps-script/Code.gs` | Backend — paste into Apps Script editor |
| `apps-script/Setup.md` | This guide |
