# Deploy to Railway — Step by Step

## What you need
- A free GitHub account (github.com)
- A free Railway account (railway.app)
- Your Anthropic API key (from console.anthropic.com)

---

## Step 1 — Put the code on GitHub

1. Go to **github.com** and sign in (or create a free account)
2. Click the **+** button (top right) → **New repository**
3. Name it `bank-analyzer`, set it to **Private**, click **Create repository**
4. On the next page, click **uploading an existing file**
5. Drag and drop ALL files from this folder:
   - `app.py`
   - `Procfile`
   - `railway.toml`
   - `requirements.txt`
   - `runtime.txt`
   - The entire `templates/` folder (drag the folder itself)
6. Click **Commit changes**

---

## Step 2 — Deploy on Railway

1. Go to **railway.app** and click **Start a New Project**
2. Sign in with your GitHub account (click "Login with GitHub")
3. Click **Deploy from GitHub repo**
4. Select your `bank-analyzer` repository
5. Railway will start building automatically — wait about 2 minutes

---

## Step 3 — Add your Anthropic API key

This is the most important step — without it the AI won't work.

1. In Railway, click on your project
2. Click **Variables** (left sidebar)
3. Click **New Variable**
4. Set:
   - **Name:** `ANTHROPIC_API_KEY`
   - **Value:** your API key (starts with `sk-ant-...`)
5. Click **Add** — Railway will automatically restart the app

---

## Step 4 — Get your live URL

1. In Railway, click **Settings** → **Networking**
2. Click **Generate Domain**
3. You'll get a URL like `https://bank-analyzer-production.up.railway.app`
4. Share this URL with your team — that's it!

---

## Monthly Cost
- Railway Hobby plan: **$5/month** flat
- That covers everything for a small team

---

## Updating the app in the future
If you ever need to change anything:
1. Edit the file on GitHub (click the file → pencil icon)
2. Commit the change
3. Railway automatically redeploys in ~2 minutes

No coding required.
