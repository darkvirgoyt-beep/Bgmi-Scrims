# 🚀 Deploy to Heroku (5 Minutes)

## Step 1: Install Heroku CLI
```bash
# On Windows/Mac/Linux
# Download from: https://devcenter.heroku.com/articles/heroku-cli
```

## Step 2: Login to Heroku
```bash
heroku login
```
This will open your browser. Login with your Heroku account.

## Step 3: Create Heroku App
```bash
cd /path/to/bgmi-scrims-repo
heroku create bgmi-scrims-virgo
```

## Step 4: Set Environment Variables
```bash
heroku config:set GOOGLE_SHEET_ID=11joNjVKYi4CrnNOBHYMX_aIer4eiKnzpg0VaWfm2WdQ
heroku config:set EMAIL_USER=imamericanyess@gmail.com
heroku config:set EMAIL_PASSWORD=bqce odep pzlh ghbs
heroku config:set RECIPIENTS=imamericanyess@gmail.com,kumarabhinav7349@gmail.com
heroku config:set ADMIN_KEY=Princeraj@1331
heroku config:set GMAIL_APP_PASSWORD=bqce odep pzlh ghbs
```

## Step 5: Deploy Code
```bash
git push heroku main
```

## Step 6: Check Logs
```bash
heroku logs --tail
```

## Step 7: Your Backend URL
```
https://bgmi-scrims-virgo.herokuapp.com
```

---

## ✅ That's It!

Your backend is now live on Heroku! 🎉

The frontend (index.html) will automatically connect to it.

---

## 🔧 Troubleshooting

**If deployment fails:**
```bash
# Check logs
heroku logs --tail

# Restart app
heroku restart

# Check config
heroku config
```

**If emails not sending:**
- Make sure Gmail App Password is correct
- Check: https://myaccount.google.com/apppasswords

**If Google Sheets not working:**
- Make sure Sheet ID is correct
- Sheet must be shared with your email
