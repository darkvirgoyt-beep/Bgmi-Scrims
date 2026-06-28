# Setup — BGMI Scrim Registration Site

Two files, three things to set up. All free, no server needed.

## 1. Backend (Apps Script) — sends the emails

1. Go to **script.google.com** → New project
2. Delete the default code, paste in everything from `apps-script-backend.gs`
3. (Optional) Make a blank Google Sheet, copy its ID from the URL, paste into `SHEET_ID` in the script — this logs every entry as a row so you can track who's confirmed
4. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Copy the **Web app URL** it gives you (looks like `https://script.google.com/macros/s/.../exec`)
6. First time it runs it'll ask you to authorize permissions (Gmail + Drive) — approve it, that's your own account

## 2. Google Sign-In — needs a Client ID

1. Go to **console.cloud.google.com** → create a project (any name)
2. **APIs & Services → OAuth consent screen** → set up as "External", add your app name + your email
3. **APIs & Services → Credentials → Create Credentials → OAuth client ID**
   - Application type: **Web application**
   - Authorized JavaScript origins: add the URL where you'll host the site (e.g. `https://darkvirgoyt-beep.github.io`)
4. Copy the **Client ID** (ends in `.apps.googleusercontent.com`)

## 3. Wire it together

Open `index.html`, find the `CONFIG` block near the bottom, and fill in:

```js
GOOGLE_CLIENT_ID: "...apps.googleusercontent.com",   // from step 2
APPS_SCRIPT_URL: "https://script.google.com/macros/s/.../exec", // from step 1
```

Everything else in `CONFIG` (entry fee, schedule, prize text, WhatsApp/YouTube links) is already filled in — edit any line there to update it without touching the rest of the file.

## 4. Host it

Easiest with your GitHub Actions habits: push `index.html` to a repo, turn on **GitHub Pages** in repo settings, done. Just make sure the Pages URL matches what you added as an "Authorized origin" in step 2.

## How payment verification works

There's no automatic "money received" check here — UPI doesn't expose that to a website without a paid gateway (Razorpay etc.) and KYC. So the flow is:

- User pays the QR, ticks "I've paid," uploads a screenshot
- That screenshot + all squad info lands in your inbox (and the sheet, if you set one up) marked **Pending**
- You manually confirm before posting the room ID in the WhatsApp group

The QR auto-regenerates every 5 minutes (new reference code each time) so a screenshotted/shared QR goes stale — but that's a UX nudge, not real fraud-proofing. Don't rely on it as your only check; eyeball each screenshot.
