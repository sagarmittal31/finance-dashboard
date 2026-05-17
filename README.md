# Personal Finance Dashboard

Interactive expense dashboard deployed via GitHub Pages.

## One-time Setup

### 1. Initialize git and push to GitHub

```bash
cd ~/Desktop/Personal\ Finances/finances-dashboard

git init
git branch -M main
git add .gitignore README.md docs/ update_dashboard.py
git commit -m "Initial commit — finance dashboard"
```

Go to https://github.com/new and create a **new empty repository** (no README, no .gitignore). Then:

```bash
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
git push -u origin main
```

### 2. Enable GitHub Pages

1. Go to your repo on GitHub → **Settings** → **Pages**
2. Under **Source**, select **Deploy from a branch**
3. Branch: `main`, Folder: `/docs`
4. Click **Save**

Your dashboard will be live at:
`https://YOUR_USERNAME.github.io/YOUR_REPO_NAME/`

---

## Updating the Dashboard

### Drop new data files into `data/`

```
finances-dashboard/
└── data/
    ├── Primary data.xlsx          ← replace with latest
    ├── Secondary data.csv         ← replace with latest
    └── Expense Categories .docx   ← replace only if rulebook changed
```

### Run the update script

```bash
cd ~/Desktop/Personal\ Finances/finances-dashboard
python update_dashboard.py
```

This will:
1. Process and categorize all expenses
2. Generate `docs/index.html`
3. Commit and push to GitHub automatically

Your dashboard URL updates within ~30 seconds.

### Options

```bash
# Preview locally without pushing to GitHub
python update_dashboard.py --no-push

# Custom commit message
python update_dashboard.py --message "Add May 2026 expenses"

# Custom file paths
python update_dashboard.py \
  --primary "data/Primary data.xlsx" \
  --secondary "data/Secondary data.csv" \
  --categories "data/Expense Categories .docx"
```

### Preview locally

After running with `--no-push`, open `docs/index.html` in your browser:

```bash
open docs/index.html
```
