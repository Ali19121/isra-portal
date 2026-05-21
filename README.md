# ISRA Schools — Online Test Portal

A professional teacher recruitment test portal hosted on GitHub Pages.

## Live URL
After deployment your portal will be at:
```
https://YOUR-GITHUB-USERNAME.github.io/isra-portal/
```

## Features
- 6 test categories (Sindhi, Maths, English — Hyderabad & Mirpurkhas)
- MCQ + Descriptive questions
- 60-minute countdown timer
- Auto-fullscreen + anti-cheat detection
- Auto-advance on MCQ answer
- AI analysis of descriptive answers
- PDF report auto-download on submission
- Excel export of candidate history
- Centralized Google Sheets backend (questions + results)
- Word document (.docx) question import
- Admin panel with password protection

## Setup (One Time)

### Step 1 — Google Apps Script
1. Go to [script.google.com](https://script.google.com)
2. Create new project → paste `apps_script.js` code
3. Run `setupSheets()` once manually
4. Copy the **Admin Secret Key** shown
5. Deploy as Web App → Anyone → copy URL

### Step 2 — Admin Panel
1. Open your portal URL
2. Click "Admin Panel" → password: `Alibaloch-IU`
3. Go to Settings → paste Apps Script URL + Admin Secret Key
4. Upload questions via Word import or manually

## Admin Password
```
Alibaloch-IU
```

## Files
- `index.html` — The complete portal (single file)
- `apps_script.js` — Google Apps Script backend code
- `README.md` — This file
