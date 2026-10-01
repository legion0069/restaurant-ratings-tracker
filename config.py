
import os
from pathlib import Path
from dotenv import load_dotenv

# Base Directories
BASE_DIR = Path(__file__).resolve().parent
if os.getenv("VERCEL"):
    DATA_DIR = Path("/tmp/data")
else:
    DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True, parents=True)

# Load .env file
ENV_FILE = BASE_DIR / ".env"
if ENV_FILE.exists():
    load_dotenv(ENV_FILE)

# Email Settings
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
EMAIL_RECIPIENT = os.getenv("EMAIL_RECIPIENT", "vsaiteja@gmail.com")

# Schedule Time
SCHEDULE_TIME = os.getenv("SCHEDULE_TIME", "09:00")

# Optional API Keys
GOOGLE_PLACES_API_KEY = os.getenv("GOOGLE_PLACES_API_KEY", "")
SERPAPI_KEY = os.getenv("SERPAPI_KEY", "")

# File Paths
EXCEL_FILE_PATH = DATA_DIR / "restaurant_ratings.xlsx"

# Restaurant Configuration
RESTAURANTS_CONFIG = {
    "kiplings": {
        "name": "Kipling's Déli & Bistro",
        "display_name": "Kiplings",
        "location": "Inorbit Mall, Madhapur, Hyderabad",
        "zomato_slug": "kiplings-deli-bistro-hitech-city",
        "swiggy_name": "Kiplings Deli and Bistro",
        "swiggy_id": 1431945,
        "lat": 17.4338,
        "lng": 78.3855,
        "google_query": "Kiplings Deli & Bistro Hyderabad",
        "google_place_id": "ChIJSUG0WACRyzsRuUFml9cpIFQ"
    },
    "casa_loco": {
        "name": "Casa Loco Express",
        "display_name": "Casa Loco Express",
        "location": "Building 9, Raheja Mindspace, Hitech City, Hyderabad",
        "zomato_slug": "casa-loco-express-hitech-city",
        "swiggy_name": "Casa loco Express",
        "swiggy_id": 1339907,
        "lat": 17.4400802,
        "lng": 78.3809632,
        "google_query": "Casa Loco Express Mindspace Hyderabad",
        "google_place_id": "ChIJX0fzg0STyzsRpHG-gcZKTlk"
    }
}
