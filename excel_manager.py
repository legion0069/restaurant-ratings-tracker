import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime, timedelta
from pathlib import Path
from config import EXCEL_FILE_PATH, BASE_DIR

# Historical seed data from user's table
HISTORICAL_SEED_DATA = [
    ["15 Sept", 116, 4.4, 153, 4.3, 28, 5.0, 63, 4.3, 46, 3.8],
    ["16 Sept", 121, 4.4, 158, 4.4, 28, 5.0, 63, 4.3, 46, 3.8],
    ["17 Sept", 121, 4.4, 164, 4.4, 28, 5.0, 65, 4.3, 46, 3.8],
    ["18 Sept", 120, 4.4, 166, 4.4, 28, 5.0, 66, 4.3, 46, 3.8],
    ["19 Sept", 121, 4.4, 171, 4.4, 28, 5.0, 67, 4.3, 46, 3.8],
    ["20 Sept", 121, 4.4, 178, 4.4, 28, 5.0, 67, 4.3, 46, 3.8],
    ["21 Sept", 123, 4.4, 187, 4.4, 28, 5.0, 67, 4.3, 46, 3.8],
    ["22 Sept", 124, 4.5, 192, 4.4, 28, 5.0, 68, 4.3, 46, 3.8],
    ["23 Sept", 125, 4.5, 194, 4.3, 28, 5.0, 68, 4.3, 46, 3.8],
    ["24 Sept", 124, 4.5, 198, 4.3, 28, 5.0, 70, 4.3, 46, 3.8],
    ["25 Sept", 129, 4.5, 200, 4.4, 28, 5.0, 71, 4.2, 47, 3.8],
    ["27 Sept", 142, 4.5, 217, 4.4, 28, 5.0, 73, 4.3, 47, 3.8],
    ["28 Sept", 143, 4.5, 222, 4.4, 28, 5.0, 73, 4.3, 48, 3.7],
    ["29 Sept", 143, 4.5, 225, 4.4, 28, 5.0, 73, 4.3, 50, 3.7],
    ["30 Sept", 143, 4.5, 229, 4.4, 28, 5.0, 73, 4.3, 52, 3.7],
    ["1 Oct", 143, 4.5, 231, 4.4, 28, 5.0, 74, 4.2, 52, 3.7],
]

def format_date_label(dt=None):
    """
    Formats a datetime object or date into uniform '1 Oct', '23 Sept' format.
    """
    if dt is None:
        dt = datetime.now()
    day = str(dt.day)
    month = "Sept" if dt.strftime("%b") == "Sep" else dt.strftime("%b")
    return f"{day} {month}"


def get_yesterday_date_label(reference_dt=None):
    """
    Returns yesterday's date label (e.g. '1 Oct' when called on Oct 2nd).
    Since the 9:00 AM daily report records the prior day's performance,
    the date recorded in the Excel tracker is yesterday's date.
    """
    ref = reference_dt if reference_dt is not None else datetime.now()
    yesterday = ref - timedelta(days=1)
    return format_date_label(yesterday)


def format_today_label(dt=None):
    """
    Backwards-compatible helper. Formats given date or today's date.
    """
    return format_date_label(dt)


def normalize_date_label(date_val):
    """
    Normalizes any datetime object or date string into uniform '1 Oct', '23 Sept' format.
    """
    if date_val is None:
        return get_yesterday_date_label()
    
    if isinstance(date_val, datetime):
        day = str(date_val.day)
        month = "Sept" if date_val.strftime("%b") == "Sep" else date_val.strftime("%b")
        return f"{day} {month}"
    
    val_str = str(date_val).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y/%m/%d"):
        try:
            dt = datetime.strptime(val_str.split(".")[0], fmt)
            day = str(dt.day)
            month = "Sept" if dt.strftime("%b") == "Sep" else dt.strftime("%b")
            return f"{day} {month}"
        except Exception:
            pass
            
    # Normalize existing strings like "16 Sep" to "16 Sept"
    if " Sep" in val_str and " Sept" not in val_str:
        val_str = val_str.replace(" Sep", " Sept")

    return val_str


def load_or_create_rows():
    """
    Loads existing data rows from Excel if file exists, or returns historical seed.
    """
    bundled_file = BASE_DIR / "data" / "restaurant_ratings.xlsx"
    source_file = EXCEL_FILE_PATH if EXCEL_FILE_PATH.exists() else (bundled_file if bundled_file.exists() else None)
    
    if not source_file:
        return list(HISTORICAL_SEED_DATA)
    
    try:
        wb = openpyxl.load_workbook(source_file, data_only=True)
        ws = wb.active
        rows = []
        # Data starts from row 4
        for r in range(4, ws.max_row + 1):
            date_val = ws.cell(row=r, column=1).value
            if not date_val:
                continue
            normalized_date = normalize_date_label(date_val)
            row_vals = [normalized_date]
            for c in range(2, 12):
                val = ws.cell(row=r, column=c).value
                row_vals.append(val)
            rows.append(row_vals)
        return rows if rows else list(HISTORICAL_SEED_DATA)
    except Exception as e:
        print(f"[Excel] Error loading existing workbook: {e}")
        return list(HISTORICAL_SEED_DATA)


def update_excel_with_ratings(ratings_dict, custom_date_str=None):
    """
    Appends or updates the tracking row in the Excel sheet and saves formatted workbook.
    By default, uses yesterday's date (e.g. 1 Oct when run on 2 Oct morning).
    """
    date_str = custom_date_str if custom_date_str else get_yesterday_date_label()
    
    k = ratings_dict.get("kiplings", {})
    c = ratings_dict.get("casa_loco", {})

    new_row = [
        date_str,
        k.get("google_count"),
        k.get("google_rating"),
        k.get("zomato_count"),
        k.get("zomato_rating"),
        c.get("google_count"),
        c.get("google_rating"),
        c.get("zomato_count"),
        c.get("zomato_rating"),
        c.get("swiggy_count"),
        c.get("swiggy_rating"),
    ]

    rows = load_or_create_rows()
    
    # Check if date_str is already in rows
    updated = False
    for idx, r in enumerate(rows):
        if str(r[0]).strip().lower() == date_str.strip().lower():
            rows[idx] = new_row
            updated = True
            break
            
    if not updated:
        rows.append(new_row)

    # Save styled Excel
    save_styled_excel(rows)
    return rows, EXCEL_FILE_PATH


# Alias for backward compatibility
update_excel_with_today_ratings = update_excel_with_ratings


def save_styled_excel(rows_data):
    """
    Generates a beautifully formatted Excel file matching the exact image structure,
    typography (Arial/Segoe UI), borders, column widths, and row heights.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daily Ratings"

    # Show grid lines
    ws.views.sheetView[0].showGridLines = True

    # Typography & Styling tokens matching reference image
    font_header_main = Font(name="Arial", size=11, bold=True, color="000000")
    font_header_platform = Font(name="Arial", size=11, bold=True, color="000000")
    font_header_metric = Font(name="Arial", size=10, bold=True, color="000000")
    font_regular = Font(name="Arial", size=10, bold=False, color="000000")
    
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    
    thin_border_side = Side(border_style="thin", color="D3D3D3")
    medium_divider_side = Side(border_style="medium", color="7F7F7F")
    
    border_header_standard = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side
    )
    border_header_bottom = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=medium_divider_side
    )
    border_data_cell = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side
    )
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_data = Alignment(horizontal="center", vertical="center")

    # 1. Header Values & Merges
    ws["A1"] = "Date"
    ws["B1"] = "Kiplings"
    ws["F1"] = "Casa Loco Express"

    ws["B2"] = "Google"
    ws["D2"] = "Zomato"
    ws["F2"] = "Google"
    ws["H2"] = "Zomato - Delivery"
    ws["J2"] = "Swiggy - Delivery"

    headers_row_3 = [
        "", # Date
        "Total no", "Rating", # Kiplings Google
        "Total no", "Rating", # Kiplings Zomato
        "Total no", "Rating", # Casa Loco Google
        "Total no", "Rating", # Casa Loco Zomato Delivery
        "Total no", "Rating", # Casa Loco Swiggy Delivery
    ]
    for col_idx in range(2, 12):
        ws.cell(row=3, column=col_idx, value=headers_row_3[col_idx-1])

    # Merge cell ranges
    merge_list = [
        "A1:A3",
        "B1:E1",
        "F1:K1",
        "B2:C2",
        "D2:E2",
        "F2:G2",
        "H2:I2",
        "J2:K2",
    ]
    for m in merge_list:
        ws.merge_cells(m)

    # Row heights
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 22

    # Apply formatting to all header cells
    for r in range(1, 4):
        for c in range(1, 12):
            cell = ws.cell(row=r, column=c)
            cell.alignment = align_center
            cell.fill = fill_white
            
            if r == 1:
                cell.font = font_header_main
            elif r == 2:
                cell.font = font_header_platform
            else:
                cell.font = font_header_metric
            
            if r == 3:
                cell.border = border_header_bottom
            else:
                cell.border = border_header_standard

    # Ensure Date merged cells have proper bounding borders
    ws["A1"].border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=medium_divider_side)
    ws["A2"].border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    ws["A3"].border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=medium_divider_side)

    # 2. Write Data Rows
    current_row = 4
    for row in rows_data:
        ws.row_dimensions[current_row].height = 20
        for col_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=current_row, column=col_idx)
            cell.font = font_regular
            cell.border = border_data_cell
            cell.fill = fill_white
            cell.alignment = align_data
            
            # Format numbers & values
            if col_idx == 1:
                cell.value = str(val) if val is not None else ""
            elif col_idx in [3, 5, 7, 9, 11]:  # Rating columns (e.g. 4.4, 5.0, 3.8)
                try:
                    if val is not None and str(val).strip() != "":
                        cell.value = float(val)
                        cell.number_format = "0.0"
                    else:
                        cell.value = ""
                except (ValueError, TypeError):
                    cell.value = val
            elif col_idx in [2, 4, 6, 8, 10]:  # Count columns (e.g. 116, 28, 63)
                try:
                    if val is not None and str(val).strip() != "":
                        cell.value = int(val)
                        cell.number_format = "#,##0"
                    else:
                        cell.value = ""
                except (ValueError, TypeError):
                    cell.value = val
            else:
                cell.value = val
                
        current_row += 1

    # 3. Adjust Column Widths to match exact proportions
    ws.column_dimensions['A'].width = 13.5
    for col_letter in ['B', 'D', 'F', 'H', 'J']:
        ws.column_dimensions[col_letter].width = 11.5
    for col_letter in ['C', 'E', 'G', 'I', 'K']:
        ws.column_dimensions[col_letter].width = 10.5

    # Save to file
    EXCEL_FILE_PATH.parent.mkdir(exist_ok=True, parents=True)
    wb.save(EXCEL_FILE_PATH)
    print(f"[Excel] Updated spreadsheet at: {EXCEL_FILE_PATH}")


if __name__ == "__main__":
    # Test update with dummy data
    test_ratings = {
        "kiplings": {"google_count": 124, "google_rating": 4.5, "zomato_count": 194, "zomato_rating": 4.3},
        "casa_loco": {"google_count": 28, "google_rating": 5.0, "zomato_count": 68, "zomato_rating": 4.3, "swiggy_count": 46, "swiggy_rating": 3.8}
    }
    rows, path = update_excel_with_today_ratings(test_ratings)
    print(f"Total rows in Excel: {len(rows)}")
