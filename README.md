# AGR Construction — Site Diary System

A web-based daily site diary management system for AGR Construction (LADOL Power Station, Contract No. LPS-003).

## Features

- Create, view, edit, and delete daily site diary entries
- Record activities, deliveries, collections, labour resources, and plant/equipment
- AM/PM time pickers for working hours
- Export entries to formatted Excel (.xlsx) with company logo header
- Print-ready PDF view (A4 landscape, single page)
- Automatic day-of-week detection from selected date

---

## Running Locally on Windows

### Prerequisites
- Python 3.10 or higher — download from [python.org](https://www.python.org/downloads/)
- pip (comes with Python)

### Setup Steps

1. **Open Command Prompt or PowerShell** and navigate to this folder:
   ```
   cd path\to\site-diary-vercel
   ```

2. **Create a virtual environment** (recommended):
   ```
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```
   pip install -r requirements.txt
   ```

4. **Run the app**:
   ```
   python main.py
   ```

5. **Open your browser** and go to: `http://localhost:5000`

The SQLite database (`site_diary.db`) will be created automatically in the `instance/` folder on first run.

---

## Deploying to Vercel

### Prerequisites
- A [Vercel account](https://vercel.com) (free tier works)
- [Vercel CLI](https://vercel.com/docs/cli) installed: `npm i -g vercel`
- A PostgreSQL database (recommended: [Neon](https://neon.tech), [Supabase](https://supabase.com), or [PlanetScale](https://planetscale.com) — all have free tiers)

### Important: Database for Vercel

Vercel's serverless environment does **not** support SQLite for persistent storage (the filesystem is reset on each deployment). You must use a cloud PostgreSQL database.

**Recommended: Neon (free tier)**
1. Sign up at [neon.tech](https://neon.tech)
2. Create a new project and database
3. Copy the connection string (it starts with `postgresql://`)

### Deployment Steps

1. **Install Vercel CLI**:
   ```
   npm i -g vercel
   ```

2. **Log in**:
   ```
   vercel login
   ```

3. **Deploy from this folder**:
   ```
   vercel
   ```
   Follow the prompts. When asked about the framework, select **Other**.

4. **Set your database environment variable** in the Vercel dashboard:
   - Go to your project → Settings → Environment Variables
   - Add: `DATABASE_URL` = your PostgreSQL connection string
   - Add: `SECRET_KEY` = any long random string

5. **Redeploy** to apply the environment variables:
   ```
   vercel --prod
   ```

### Environment Variables

| Variable | Description | Example |
|---|---|---|
| `DATABASE_URL` | PostgreSQL connection string | `postgresql://user:pass@host/dbname` |
| `SECRET_KEY` | Flask secret key for sessions | `mysupersecretkey123` |

---

## Project Structure

```
site-diary-vercel/
├── app.py              # Main Flask application
├── main.py             # Entry point (runs the app)
├── requirements.txt    # Python dependencies
├── vercel.json         # Vercel deployment configuration
├── static/
│   ├── css/style.css   # Stylesheet
│   ├── js/main.js      # Frontend JavaScript
│   └── images/         # Logo and images
├── templates/
│   ├── base.html       # Base layout template
│   ├── index.html      # Entry list page
│   ├── form.html       # New/edit entry form
│   ├── view.html       # Entry detail view
│   └── print.html      # Print/PDF template
└── instance/
    └── site_diary.db   # SQLite database (local only)
```

---

## Tech Stack

- **Backend**: Python / Flask
- **Database**: SQLite (local) / PostgreSQL (Vercel)
- **ORM**: Flask-SQLAlchemy
- **Excel Export**: openpyxl + Pillow
- **Frontend**: Vanilla HTML, CSS, JavaScript
