# 🚀 PPT Generator - Hosting Guide

## Option 1: Railway.app (RECOMMENDED - Easiest)

### Steps:
1. Go to https://railway.app and sign up (use GitHub)
2. Click "New Project" → "Deploy from GitHub repo"
3. Upload your `report_engine` folder to GitHub first:
   ```
   cd report_engine
   git init
   git add .
   git commit -m "Initial commit"
   git remote add origin https://github.com/YOUR_USERNAME/report_engine.git
   git push -u origin main
   ```
4. Select your repo in Railway
5. Railway will auto-detect Python and deploy
6. Go to Settings → Networking → Generate Domain
7. Your API will be at: `https://your-app.railway.app/generate-ppt`

### Cost: FREE (500 hours/month)

---

## Option 2: Render.com (Also Easy)

### Steps:
1. Go to https://render.com and sign up
2. Click "New" → "Web Service"
3. Connect your GitHub repo
4. Settings:
   - Environment: Python
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn api_server:app`
5. Deploy!
6. Your URL: `https://your-app.onrender.com/generate-ppt`

### Cost: FREE (spins down after 15 min inactivity, 750 hours/month)

---

## Option 3: PythonAnywhere (Python-focused)

### Steps:
1. Go to https://www.pythonanywhere.com and sign up
2. Go to "Web" tab → Add new web app
3. Choose Flask, Python 3.11
4. Upload your files via "Files" tab
5. Edit WSGI configuration file to point to your app
6. Reload web app

### Cost: FREE (limited, your-username.pythonanywhere.com)

---

## Option 4: DigitalOcean/VPS (More Control)

For $4-6/month, get a VPS and run:
```bash
# On your VPS
git clone https://github.com/YOUR_USERNAME/report_engine.git
cd report_engine
pip install -r requirements.txt
gunicorn api_server:app --bind 0.0.0.0:5000 --daemon
```

Use nginx as reverse proxy for HTTPS.

---

## 🔧 Important: Update n8n After Hosting

Once deployed, update your n8n HTTP Request node URL from:
```
https://xxx.loca.lt/generate-ppt
```
To:
```
https://your-app.railway.app/generate-ppt
```

---

## Files Required for Deployment:
- api_server.py ✓
- ppt_generator.py ✓
- master_template.pptx ✓
- requirements.txt ✓
- Procfile ✓
- runtime.txt ✓

All files are ready! Just push to GitHub and deploy.
