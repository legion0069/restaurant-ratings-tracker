# 🍽️ Daily Restaurant Ratings Automation & Mailer

An automated, end-to-end tracking and reporting pipeline that fetches live ratings and total review counts for **Kipling's Déli & Bistro** and **Casa Loco Express** across **Google**, **Zomato**, and **Swiggy**, updates a formatted Excel spreadsheet, and sends a daily email report with the updated workbook attached.

---

## ✨ Features

- **Live Multi-Platform Scraping**: Automatically extracts rating stars and total review counts from **Google Maps**, **Zomato**, and **Swiggy**.
- **Automated Excel Spreadsheet**: Maintains a multi-level structured Excel file (`data/restaurant_ratings.xlsx`) using `openpyxl`.
- **Responsive HTML Daily Email**: Generates an email report featuring:
  - Color-coded multi-tier table matching your spreadsheet format.
  - **Day-over-day change badges** (e.g. `(+5)`, `(+0.1)`) highlighting new reviews and rating changes.
  - Formatted `.xlsx` file attached to every daily email.
- **Resilient Email Delivery**: Built-in dual-protocol support (STARTTLS on port `587` & SSL on port `465`) with automatic retries.
- **Multiple Scheduling Methods**: Supports local Windows Task Scheduler, continuous Python background loop, and 100% serverless GitHub Actions.

---

## 📊 Spreadsheet & Report Layout

The Excel workbook (`data/restaurant_ratings.xlsx`) and the daily HTML email match the following structure:

| Date | Kiplings (Google: Total no, Rating \| Zomato: Total no, Rating) | Casa Loco Express (Google: Total no, Rating \| Zomato - Delivery: Total no, Rating \| Swiggy - Delivery: Total no, Rating) |
| :--- | :--- | :--- |

### Detailed Column Mapping:

```
+-----------+-----------------------------+-------------------------------------------------------+
|   Date    |           Kiplings          |                   Casa Loco Express                   |
|           +--------------+--------------+--------------+--------------------+-------------------+
|           |    Google    |    Zomato    |    Google    | Zomato - Delivery  | Swiggy - Delivery |
|           +-------+------+-------+------+-------+------+--------+-----------+--------+----------+
|           | Total | Rate | Total | Rate | Total | Rate | Total  |   Rate    | Total  |   Rate   |
+-----------+-------+------+-------+------+-------+------+--------+-----------+--------+----------+
| 22 Sept   |  124  | 4.5  |  192  | 4.4  |  28   | 5.0  |   68   |    4.3    |   46   |   3.8    |
| Today     |  125  | 4.5  |  229  | 4.4  |  28   | 5.0  |   73   |    4.3    |   52   |   3.7    |
+-----------+-------+------+-------+------+-------+------+--------+-----------+--------+----------+
```

---

## 🚀 Quick Start Guide

### 1. Prerequisites & Installation

Ensure you have **Python 3.10+** installed:

```bash
# Clone or navigate to the project directory
cd restaurant-ratings-tracker

# Install required dependencies
pip install -r requirements.txt
```

### 2. Configure Environment (`.env`)

Create or update your `.env` file in the root directory:

```env
# Email Settings (Gmail SMTP)
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=your_email@gmail.com
SMTP_PASSWORD=your_16_character_app_password
EMAIL_RECIPIENT=vsaiteja@isthara.com

# Schedule Time (24-hour format HH:MM, default 09:00 AM)
SCHEDULE_TIME=09:00

# Optional API Keys (Leave blank for default web scrapers)
GOOGLE_PLACES_API_KEY=
SERPAPI_KEY=
```

> **🔑 How to generate a Gmail App Password:**
> 1. Go to your [Google Account Security](https://myaccount.google.com/security) page.
> 2. Ensure **2-Step Verification** is enabled.
> 3. Go to [App Passwords](https://myaccount.google.com/apppasswords).
> 4. Create a new entry (e.g., `RatingsTracker`).
> 5. Copy the 16-character generated code and paste it as `SMTP_PASSWORD` in `.env`.

---

## 🏃 Running the Automation

### 1. Immediate Run (Fetch + Update Excel + Send Email)
Executes the full pipeline immediately:
```bash
python main.py --run-now
```

### 2. Dry Run (Fetch + Update Excel Only)
Fetches live ratings and updates the Excel file without sending an email:
```bash
python main.py --dry-run
```

### 3. Continuous Scheduler
Runs an active scheduling loop in your terminal that triggers every day at `SCHEDULE_TIME` (default: 09:00 AM):
```bash
python main.py --schedule
```

---

## ⏰ Automated Daily Scheduling (09:00 AM IST)

### 🌟 Option A: GitHub Actions (100% Free Cloud Automation — Recommended)

The repository includes a ready-to-run GitHub Actions workflow ([`.github/workflows/daily_ratings.yml`](file:///.github/workflows/daily_ratings.yml)) that runs automatically every morning at **09:00 AM IST (03:30 UTC)**.

**What it does every morning:**
1. Spins up a cloud runner and installs Python dependencies.
2. Scrapes live ratings from **Google Maps (SerpAPI)**, **Zomato**, and **Swiggy**.
3. Updates and styles [`data/restaurant_ratings.xlsx`](file:///data/restaurant_ratings.xlsx) and commits it back to your GitHub repo.
4. Sends the daily HTML report with the attached Excel file to **`vsaiteja@isthara.com`**.

**GitHub Setup:**
1. In your GitHub repository, go to **Settings ➔ Secrets and variables ➔ Actions**.
2. Add these repository secrets:
   - `SMTP_USER`: `tejaverukonda@gmail.com`
   - `SMTP_PASSWORD`: `btmupxjftvrotabb`
   - `EMAIL_RECIPIENT`: `vsaiteja@isthara.com`
   - `SERPAPI_KEY`: `6ad40cf887798c107fa84025951df57bf3e919b278cc68be6996196bf85986cb`
3. Everything is now 100% automated! You can also trigger a manual run anytime from the **Actions** tab by clicking **"Run workflow"**.

---

### 💻 Option B: Windows Task Scheduler (Local PC)

If you prefer running locally on your computer:
```powershell
.\setup_windows_task.ps1
```
> Registers a native Windows background task named `RestaurantRatingsDailyMailer` that runs `python main.py --run-now` every morning at **09:00 AM IST**, even when no terminal window is open.

---

### Option C: Windows Task Scheduler (Local PC)

Run the automated setup script to register a native Windows background task:
```powershell
.\setup_windows_task.ps1
```
> Registers a Windows Scheduled Task named `RestaurantRatingsDailyMailer` that triggers `python main.py --run-now` every morning at **09:00 AM IST**.

---

## 📁 Project Structure

```
restaurant-ratings-tracker/
├── .env                              # Local configuration & credentials (git-ignored)
├── .env.example                      # Configuration template
├── config.py                         # Application configuration & restaurant targets
├── scraper.py                        # Web scrapers for Google, Zomato & Swiggy
├── excel_manager.py                  # OpenPyXL workbook manager & styling engine
├── email_notifier.py                 # Responsive HTML report builder & SMTP mailer
├── main.py                           # CLI entry point and scheduler
├── requirements.txt                  # Python package dependencies
├── setup_windows_task.bat            # One-click Windows Task Scheduler setup (Batch)
├── setup_windows_task.ps1            # One-click Windows Task Scheduler setup (PowerShell)
├── .github/
│   └── workflows/
│       └── daily_ratings.yml         # Cloud GitHub Actions daily workflow
└── data/
    └── restaurant_ratings.xlsx       # Multi-day formatted tracking spreadsheet
```

---

## ⚙️ Configuration Reference (`.env`)

| Variable | Default | Description |
| :--- | :--- | :--- |
| `SMTP_HOST` | `smtp.gmail.com` | SMTP server address |
| `SMTP_PORT` | `587` | SMTP port (`587` for STARTTLS, `465` for SSL) |
| `SMTP_USER` | - | Sender email address |
| `SMTP_PASSWORD` | - | 16-character Gmail App Password |
| `EMAIL_RECIPIENT` | `vsaiteja@isthara.com` | Recipient email address for daily reports |
| `SCHEDULE_TIME` | `09:00` | Local time (24h `HH:MM`) for scheduled daily run |
| `SERPAPI_KEY` | _(required for Google live sync)_ | SerpAPI key for live Google Maps ratings & reviews scraping |
| `GOOGLE_PLACES_API_KEY` | _(optional alternative)_ | Google Places API key if using Google Cloud directly |

---

### 🌐 Live Platform Integrations

- **Google Maps**: Scraped in real-time using **SerpAPI Google Maps Engine** (`engine=google_maps`).
- **Zomato**: Scraped dynamically via JSON state extraction for dining & delivery metrics.
- **Swiggy**: Scraped dynamically via Swiggy DAPI search endpoint.
