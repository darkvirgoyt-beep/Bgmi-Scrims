# Setup — BGMI Scrim Registration Site

Two files, three things to set up. All free, no server needed.

## 1. Backend (Apps Script) — sends emails AND powers live settings

1. Make a blank Google Sheet (just a new empty spreadsheet). Copy its ID from
   the URL — the long string between `/d/` and `/edit`.
2. Go to **script.google.com** → New project
3. Delete the default code, paste in everything from `apps-script-backend.gs`
4. Paste your Sheet ID into `SHEET_ID` in the script
5. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
6. Copy the **Web app URL** it gives you (ends in `/exec`)
7. First run will ask to authorize Gmail + Drive + Sheets access — that's your
   own account, approve it
8. Open the Sheet — a **"Settings"** tab appears automatically the first time
   the site loads (or run `setupSettingsSheet` once from the Apps Script
   editor's Run button to create it immediately)

## 2. Google Sign-In — needs a Client ID

1. console.cloud.google.com → create a project
2. APIs & Services → OAuth consent screen → External → fill basic info
3. APIs & Services → Credentials → Create Credentials → OAuth client ID
   - Web application
   - Authorized JavaScript origins: your GitHub Pages URL (e.g.
     `https://darkvirgoyt-beep.github.io`)
4. Copy the Client ID

## 3. Wire it together

Open `index.html`, find `CONFIG` near the bottom, fill in:

```js
GOOGLE_CLIENT_ID: "...apps.googleusercontent.com",
APPS_SCRIPT_URL: "https://script.google.com/macros/s/.../exec",
```

## 4. Host it

Push the files to a GitHub repo → Settings → Pages → branch `main` → Save.
Your link stays the same forever — this is the one link you send to the
WhatsApp group every time.

---

## Round-based sheets (one per map per date)

Every time the **active map** changes — by editing the Settings sheet or
tapping a map button in the admin panel — new registrations go into a
fresh sheet named `<Map>_<Date>`, e.g. `Livik_2026-06-30`. This is
automatic: the moment you switch maps, the player count for that round
starts at 0, and old rounds stay exactly as they were, untouched.

This means: if you run Erangel today and Livik tomorrow, you get two
separate sheets, never mixed. If you run Erangel twice in one day, both
registrations land in the same `Erangel_<today>` sheet (treated as one
continuous round). Old data from before this update (the flat
**Registrations** tab) is left alone — it just stops being written to.

## Admin Panel — your control room

Open `yoursite.com/admin.html` (same repo, same Pages hosting — not
linked anywhere on the public site on purpose). Log in with the password
you set in the Sheet's **adminPassword** row (default is a placeholder —
change it before relying on this).

From there, with no spreadsheet digging required:
- See the live capacity bar for whichever map is currently active
- Tap **Erangel / Livik / Miramar** to switch the active map instantly —
  this is also how you "select" a map before sending the link out; only
  the active one is open for registration, the other two are simply not
  what the public site is pointing at
- Type the **Room ID** and **Lobby Time**, tick the broadcast box, and
  every team marked Approved gets it emailed instantly
- Scroll through every registered team for the current round — name,
  UIDs, WhatsApp, payment screenshot link, and a one-tap **Approve**
  button per team

Security note: the password is checked on your own Apps Script backend,
not just hidden in the page, so it's a real gate. But there's no full
login system behind it, so don't post the admin.html link publicly —
treat it like a spare key, not a public page.

## Changing match details (schedule, prizes, links, entry fee, caps)

**You never touch code or GitHub for this.** Open the Google Sheet, tap
the **Settings** tab, and edit the **Value** column next to whichever
row you want to change:

| Key | What it controls |
|---|---|
| `mode` | "Map · Mode" descriptor (shown in emails) |
| `schedule` | match schedule text |
| `winPrize` | "Chicken Dinner" chip |
| `mvpPrize` | "MVP Bonus" chip |
| `liveStream` | the YouTube link in the header / success screen |
| `whatsappGroup` | the WhatsApp link in the header / success screen |
| `entryFee` | the ₹ amount and the UPI QR code |
| `upiId` | the UPI ID the QR points to |
| `upiPayeeName` | the name shown in the UPI app when scanned |
| `activeMap` | which map is currently open (`erangel` / `livik` / `miramar`) — same as tapping a button in admin.html |
| `erangelMax` / `livikMax` / `miramarMax` | player cap per map |
| `roomId` / `roomTime` | private — only ever emailed, never shown publicly |
| `adminPassword` | the password for admin.html — change this first |

Save the cell, and anyone who opens (or refreshes) your site link sees
the new values within a second or two — same link, no re-deploy.

## How payment verification actually works

There's no official "money received" check — UPI doesn't expose that to
a website without a paid gateway (Razorpay etc.) and KYC. Two layers
here instead:

1. **The FamPay notification hack (MacroDroid):** when a real payment
   notification fires on your phone, it's relayed to the backend, which
   tries to match it to a pending team by transaction ref or by amount.
   If exactly one match is found, it auto-approves and emails the
   captain. If it's ambiguous, it emails you instead of guessing.
2. **Manual fallback:** every team also uploads a payment screenshot
   that's checked on-device for the UPI ID before they can even submit,
   then saved to your Drive — visible from the admin panel or the round
   sheet, for whenever you want to eyeball one yourself.

Both are best-effort, not bulletproof — worth a periodic glance at the
admin panel, especially early on, to make sure the auto-matching is
behaving the way you expect.
