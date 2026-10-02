import smtplib
import ssl
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
from config import SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, EMAIL_RECIPIENT

def generate_html_table(rows_data):
    """
    Builds clean, responsive HTML table matching the spreadsheet structure.
    """
    last_row = rows_data[-1]
    prev_row = rows_data[-2] if len(rows_data) > 1 else None
    
    rows_html = ""
    for r_idx, r in enumerate(rows_data):
        is_latest = (r_idx == len(rows_data) - 1)
        row_bg = "#EFF6FF" if is_latest else ("#FFFFFF" if r_idx % 2 == 0 else "#F8FAFC")
        font_weight = "600" if is_latest else "400"
        
        row_cells = f"<td style='padding: 9px 10px; border: 1px solid #CBD5E1; text-align: center; font-weight: {font_weight};'>{r[0]}</td>"
        for c_idx in range(1, len(r)):
            val = r[c_idx]
            if val is None:
                formatted_val = "-"
            elif isinstance(val, float):
                formatted_val = f"{val:.1f}"
            else:
                formatted_val = f"{val}"
            
            diff_badge = ""
            if is_latest and prev_row and isinstance(val, (int, float)) and isinstance(prev_row[c_idx], (int, float)):
                diff = val - prev_row[c_idx]
                if diff > 0:
                    diff_badge = f"<span style='color: #16A34A; font-size: 10px; font-weight: bold; margin-left: 2px;'> (+{diff if isinstance(val, int) else f'{diff:.1f}'})</span>"
                elif diff < 0:
                    diff_badge = f"<span style='color: #DC2626; font-size: 10px; font-weight: bold; margin-left: 2px;'> ({diff if isinstance(val, int) else f'{diff:.1f}'})</span>"
            
            row_cells += f"<td style='padding: 9px 10px; border: 1px solid #CBD5E1; text-align: center; font-weight: {font_weight};'>{formatted_val}{diff_badge}</td>"
        
        rows_html += f"<tr style='background-color: {row_bg};'>{row_cells}</tr>"

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{
                font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
                color: #1E293B;
                background-color: #F1F5F9;
                margin: 0;
                padding: 16px;
            }}
            .card {{
                max-width: 920px;
                margin: 0 auto;
                background: #FFFFFF;
                border-radius: 8px;
                overflow: hidden;
                box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
                border: 1px solid #E2E8F0;
            }}
            .header {{
                background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
                color: #FFFFFF;
                padding: 20px 24px;
                text-align: left;
            }}
            .header h1 {{
                margin: 0;
                font-size: 20px;
                font-weight: 700;
                letter-spacing: -0.02em;
            }}
            .header p {{
                margin: 4px 0 0 0;
                opacity: 0.9;
                font-size: 13px;
            }}
            .body-content {{
                padding: 20px 24px;
            }}
            .table-container {{
                overflow-x: auto;
                margin-top: 12px;
                border-radius: 6px;
                border: 1px solid #CBD5E1;
            }}
            table {{
                width: 100%;
                border-collapse: collapse;
                font-size: 12px;
                text-align: center;
            }}
            th {{
                padding: 8px 6px;
                border: 1px solid #CBD5E1;
                font-weight: 600;
                color: #0F172A;
            }}
            .footer {{
                background-color: #F8FAFC;
                padding: 14px 24px;
                border-top: 1px solid #E2E8F0;
                font-size: 11px;
                color: #64748B;
                text-align: center;
            }}
        </style>
    </head>
    <body>
        <div class="card">
            <div class="header">
                <h1>Daily Restaurant Ratings Report</h1>
                <p>Kipling's Déli &amp; Bistro &bull; Casa Loco Express | Data for {last_row[0]}</p>
            </div>
            <div class="body-content">
                <p style="font-size: 13px; color: #475569; margin: 0 0 14px 0;">
                    Hello! Here is your automated daily ratings summary for <strong>{last_row[0]}</strong>. The updated Excel spreadsheet is attached.
                </p>
                <div class="table-container">
                    <table>
                        <thead>
                            <tr>
                                <th rowspan="3" style="vertical-align: middle; background-color: #F8FAFC; width: 75px;">Date</th>
                                <th colspan="4" style="background-color: #EFF6FF; border-bottom: 2px solid #3B82F6;">Kiplings</th>
                                <th colspan="6" style="background-color: #FFF7ED; border-bottom: 2px solid #F97316;">Casa Loco Express</th>
                            </tr>
                            <tr>
                                <th colspan="2" style="background-color: #DBEAFE;">Google</th>
                                <th colspan="2" style="background-color: #DBEAFE;">Zomato</th>
                                <th colspan="2" style="background-color: #FFEDD5;">Google</th>
                                <th colspan="2" style="background-color: #FFEDD5;">Zomato - Delivery</th>
                                <th colspan="2" style="background-color: #FFEDD5;">Swiggy - Delivery</th>
                            </tr>
                            <tr style="background-color: #F8FAFC; font-size: 11px;">
                                <th>Total no</th><th>Rating</th>
                                <th>Total no</th><th>Rating</th>
                                <th>Total no</th><th>Rating</th>
                                <th>Total no</th><th>Rating</th>
                                <th>Total no</th><th>Rating</th>
                            </tr>
                        </thead>
                        <tbody>
                            {rows_html}
                        </tbody>
                    </table>
                </div>
            </div>
            <div class="footer">
                Automated Daily Ratings Tracker &bull; Delivered every morning at 9:00 AM IST to {EMAIL_RECIPIENT}
            </div>
        </div>
    </body>
    </html>
    """
    return html


def send_ratings_email(rows_data, excel_path, recipient=None):
    """
    Sends the HTML summary email with the Excel attachment to one or multiple recipients.
    """
    recipient_str = recipient or EMAIL_RECIPIENT
    recipient_list = [e.strip() for e in re.split(r"[,;]+", recipient_str) if e.strip()]
    to_header = ", ".join(recipient_list)

    today_label = rows_data[-1][0] if rows_data else "Today"
    subject = f"Daily Ratings Report - Kiplings & Casa Loco Express ({today_label})"

    # Check if SMTP credentials are provided
    if not SMTP_USER or not SMTP_PASSWORD:
        print("\n[Email Warning] SMTP_USER or SMTP_PASSWORD is not configured in .env!")
        print(f"[Email Preview] Would send email to: {to_header}")
        print(f"[Email Preview] Subject: {subject}")
        print(f"[Email Preview] Attachment: {excel_path}")
        print("To enable live email sending, add your Gmail App Password to the .env file.")
        return False

    msg = MIMEMultipart()
    msg['From'] = f"Restaurant Tracker <{SMTP_USER}>"
    msg['To'] = to_header
    msg['Subject'] = subject

    # Attach HTML body
    html_content = generate_html_table(rows_data)
    msg.attach(MIMEText(html_content, 'html', 'utf-8'))

    # Attach Excel file
    excel_file = Path(excel_path)
    if excel_file.exists():
        with open(excel_file, 'rb') as f:
            part = MIMEApplication(f.read(), Name=excel_file.name)
            part['Content-Disposition'] = f'attachment; filename="{excel_file.name}"'
            msg.attach(part)

    # Send via SMTP
    attempts = [
        (SMTP_PORT, SMTP_PORT == 465),
        (465 if SMTP_PORT != 465 else 587, SMTP_PORT != 465)
    ]

    for port, use_ssl in attempts:
        try:
            protocol = "SSL" if use_ssl else "STARTTLS"
            print(f"[Email] Connecting to {SMTP_HOST}:{port} using {protocol}...")
            if use_ssl:
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL(SMTP_HOST, port, context=context, timeout=25)
            else:
                server = smtplib.SMTP(SMTP_HOST, port, timeout=25)
                server.starttls()
            
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg, to_addrs=recipient_list)
            server.quit()
            print(f"[Email SUCCESS] Successfully sent daily ratings email to {to_header}!")
            return True
        except Exception as e:
            print(f"[Email Warning] Attempt on port {port} ({'SSL' if use_ssl else 'STARTTLS'}) failed: {e}")

    print(f"[Email ERROR] Could not deliver email via configured SMTP.")
    return False


if __name__ == "__main__":
    from excel_manager import load_or_create_rows
    from config import EXCEL_FILE_PATH
    rows = load_or_create_rows()
    send_ratings_email(rows, EXCEL_FILE_PATH)
