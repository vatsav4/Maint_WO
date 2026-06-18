# MHCV Maintenance Calendar — Flask App
## Tata Motors, Dharwad | Chassis Assembly Line

---

## Folder Structure
```
maintenance_app/
├── app.py               ← Flask backend (all routes)
├── init_db.py           ← Run ONCE to create the database
├── maintenance_db.xlsx  ← Excel database (auto-created)
├── README.md
└── templates/
    ├── base.html        ← Sidebar + layout
    ├── dashboard.html   ← Home page
    ← calendar.html     ← Monthly task view + status update
    ├── equipment.html   ← Equipment CRUD
    └── tbm_cbm.html     ← TBM / CBM CRUD (shared template)
```

---

## Setup (One Time)

### 1. Install Python dependencies
```bash
pip install flask openpyxl pandas xlsxwriter
```

### 2. Initialize the database
```bash
cd maintenance_app
python init_db.py
```
This creates `maintenance_db.xlsx` with all your equipment, TBM and CBM data pre-loaded.

### 3. Run the app
```bash
python app.py
```

### 4. Open in browser
```
http://localhost:5000
```

---

## Features

| Feature | How |
|---|---|
| View current month tasks | Dashboard → "Open [Month] Tasks" |
| Mark task as Done/Pending/Deferred | Calendar view → click the status badge (cycles through) |
| Add remarks/observations | Calendar view → ✏️ Note button |
| Add new equipment | Equipment → ＋ Add Equipment |
| Edit equipment frequencies | Equipment → ✏️ Edit |
| Remove equipment | Equipment → 🗑 (soft delete, preserves history) |
| Add/edit TBM task | TBM Tasks page |
| Add/edit CBM activity | CBM Activities page |
| Export current month to Excel | Calendar → 📥 Export |
| Export full year to Excel | Dashboard → Export Year |

---

## Month Index Reference
| Index | Month |
|---|---|
| 0 | APR |
| 1 | MAY |
| 2 | JUN |
| 3 | JUL |
| 4 | AUG |
| 5 | SEP |
| 6 | OCT |
| 7 | NOV |
| 8 | DEC |
| 9 | JAN |
| 10 | FEB |
| 11 | MAR |

---

## Database: maintenance_db.xlsx

The Excel file has 5 sheets:
- **Equipment** — All 55 equipment with PM types and scheduling info
- **TBM** — 13 time-based maintenance tasks
- **CBM** — 6 condition-based monitoring activities
- **Activity_Log** — Every status update made in the app
- **FY_Config** — Current FY setting

You can open this file directly in Excel to view/edit data manually too.

---

## Auto-start on Windows (optional)
Create a `.bat` file:
```bat
@echo off
cd /d C:\path\to\maintenance_app
python app.py
```
Add it to Windows Startup folder: `shell:startup`
