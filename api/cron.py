from http.server import BaseHTTPRequestHandler
import json
import os
import sys

# Ensure root directory is in sys.path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from main import execute_daily_job

class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        """
        Endpoint invoked by Vercel Cron or manual browser/curl ping.
        """
        try:
            print("[Vercel Cron] Starting Daily Ratings execution...")
            execute_daily_job(dry_run=False)
            
            self.send_response(200)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            response = {
                "status": "success",
                "message": "Daily restaurant ratings fetched and email successfully sent."
            }
            self.wfile.write(json.dumps(response).encode('utf-8'))
        except Exception as e:
            self.send_response(500)
            self.send_header('Content-type', 'application/json')
            self.end_headers()
            response = {
                "status": "error",
                "error": str(e)
            }
            self.wfile.write(json.dumps(response).encode('utf-8'))
