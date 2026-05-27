# ADS Assignment Document System

A web-based tool that generates personalized academic assignments from a DOCX template. Features a stunning 3D space background with iOS-style design.

## ✨ Features

- **3D Space Background** — Animated stars, floating crystals, and shooting stars using Three.js
- **6 Color Themes** — Default, Sunset, Emerald, Midnight, Rose, Ocean
- **iOS Apple Design** — SF Pro font stack, glassmorphism cards
- **Document Customization** — Enter your Name, PID, and Student Code
- **Preview** — See the generated document in your browser
- **Download** — Export as DOCX (preserves original formatting) or PDF

## 🚀 Local Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Start the app
python app.py

# 3. Open in browser
open http://localhost:5000
```

## 📦 Deploy to Render

### Step 1: Push to GitHub

```bash
# Create a new repo on GitHub, then:
git remote add origin https://github.com/YOUR_USERNAME/ads-assignment-system.git
git branch -M main
git push -u origin main
```

### Step 2: Deploy on Render

1. Go to [render.com](https://render.com) and sign in
2. Click **"New +"** → **"Web Service"**
3. Connect your GitHub repository
4. Configure:
   - **Name**: `ads-assignment-system`
   - **Environment**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
5. Click **"Create Web Service"**
6. Wait 2-3 minutes for the build & deploy
7. Your live URL will be: `https://ads-assignment-system.onrender.com`

### ⚠️ Important

- The `FINAL_SATS_ASS.docx` template file **must be committed to git** (it's the source document)
- PDF generation uses the server's available fonts (Arial on Windows, DejaVu Sans on Linux)
- For best PDF results, the DOCX download preserves all original formatting

## 📁 Project Structure

```
├── app.py                  # Flask backend
├── templates/
│   └── index.html          # Frontend with 3D background
├── FINAL_SATS_ASS.docx     # DOCX template
├── requirements.txt        # Python dependencies
├── Procfile                # Render deployment config
├── .gitignore
└── README.md
```

## 🛠️ Tech Stack

- **Backend**: Python, Flask, python-docx, fpdf2
- **Frontend**: Three.js, HTML5, CSS3 (iOS design)
- **Deployment**: Render, Gunicorn
