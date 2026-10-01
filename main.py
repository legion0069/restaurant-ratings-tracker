import argparse
import sys
import time
import schedule
from datetime import datetime
from scraper import fetch_all_ratings
from excel_manager import update_excel_with_today_ratings, format_today_label
from email_notifier import send_ratings_email
from config import SCHEDULE_TIME, EMAIL_RECIPIENT

def execute_daily_job(dry_run=False):
    """
    Main job that fetches ratings, updates the Excel file, and dispatches the email.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"\n[{timestamp}] Starting Daily Ratings Job...")

    try:
        # Step 1: Scrape ratings
        ratings = fetch_all_ratings()

        # Step 2: Update Excel
        rows_data, excel_path = update_excel_with_today_ratings(ratings)
        print(f"[{timestamp}] Excel updated with today's record ({len(rows_data)} total days tracked).")

        # Step 3: Dispatch Email
        if dry_run:
            print(f"[{timestamp}] Dry run complete. Excel saved at {excel_path}. Email skipped.")
        else:
            success = send_ratings_email(rows_data, excel_path)
            if success:
                print(f"[{timestamp}] Daily report email successfully sent to {EMAIL_RECIPIENT}!")
            else:
                print(f"[{timestamp}] Email not dispatched (see configuration notes above).")

    except Exception as e:
        print(f"[{timestamp}] ERROR in daily job execution: {e}")


def run_scheduler():
    """
    Runs the scheduling loop every day at SCHEDULE_TIME (09:00 AM).
    """
    print("==================================================")
    print(" Restaurant Ratings Automation Scheduler Active")
    print("==================================================")
    print(f"• Scheduled Time: Every day at {SCHEDULE_TIME} (Local System Time)")
    print(f"• Recipient: {EMAIL_RECIPIENT}")
    print("• Press Ctrl+C to stop scheduler.\n")

    # Schedule the job
    schedule.every().day.at(SCHEDULE_TIME).do(execute_daily_job)

    while True:
        try:
            schedule.run_pending()
            time.sleep(30)
        except KeyboardInterrupt:
            print("\nScheduler stopped by user.")
            break
        except Exception as e:
            print(f"Scheduler loop error: {e}")
            time.sleep(30)


def main():
    parser = argparse.ArgumentParser(description="Automated Daily Restaurant Ratings Tracker & Mailer")
    parser.add_argument("--run-now", action="store_true", help="Run the automation job once immediately")
    parser.add_argument("--schedule", action="store_true", help="Start the daily background scheduler (default 9:00 AM)")
    parser.add_argument("--dry-run", action="store_true", help="Fetch ratings and update Excel without sending email")
    
    args = parser.parse_args()

    if args.dry_run:
        execute_daily_job(dry_run=True)
    elif args.run_now:
        execute_daily_job(dry_run=False)
    elif args.schedule:
        run_scheduler()
    else:
        # Default behavior if no args provided: run once now and show instructions
        print("No flag provided. Running immediate update...\n")
        execute_daily_job(dry_run=False)
        print("\nTip: Run with `python main.py --schedule` to keep the scheduler running daily at 9:00 AM.")


if __name__ == "__main__":
    main()
