import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
from pathlib import Path
from config import EXCEL_FILE_PATH, BASE_DIR

# Historical seed data from user's image
HISTORICAL_SEED_DATA = [
    ["15 Sept", 116, 4.4, 153, 4.3, 28, 5.0, 63, 4.3, 46, 3.8],
    ["16 Sept", 121, 4.4, 158, 4.4, 28, 5.0, 63, 4.3, 46, 3.8],
    ["17 Sept", 121, 4.4, 164, 4.4, 28, 5.0, 65, 4.3, 46, 3.8],
    ["18 Sept", 120, 4.4, 166, 4.4, 28, 5.0, 66, 4.3, 46, 3.8],
    ["19 Sept", 121, 4.4, 171, 4.4, 28, 5.0, 67, 4.3, 46, 3.8],
    ["20 Sept", 121, 4.4, 178, 4.4, 28, 5.0, 67, 4.3, 46, 3.8],
    ["21 Sept", 123, 4.4, 187, 4.4, 28, 5.0, 67, 4.3, 46, 3.8],
    ["22 Sept", 124, 4.5, 192, 4.4, 28, 5.0, 68, 4.3, 46, 3.8],
]

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
            row_vals = [str(date_val)]
            for c in range(2, 12):
                val = ws.cell(row=r, column=c).value
                row_vals.append(val)
            rows.append(row_vals)
        return rows if rows else list(HISTORICAL_SEED_DATA)
    except Exception as e:
        print(f"[Excel] Error loading existing workbook: {e}")
        return list(HISTORICAL_SEED_DATA)


def format_today_label(dt=None):
    """
    Formats date like '24 Sept', '15 Oct', etc.
    """
    if dt is None:
        dt = datetime.now()
    day = dt.strftime("%d").lstrip("0")
    month = dt.strftime("%b")
    return f"{day} {month}"


def update_excel_with_today_ratings(ratings_dict, custom_date_str=None):
    """
    Appends or updates today's row in the Excel sheet and saves formatted workbook.
    """
    date_str = custom_date_str if custom_date_str else format_today_label()
    
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


def save_styled_excel(rows_data):
    """
    Generates a beautifully formatted Excel file matching the exact image structure.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Daily Ratings"

    # Show grid lines
    ws.views.sheetView[0].showGridLines = True

    # Styling fonts & borders
    font_header_title = Font(name="Segoe UI", size=11, bold=True, color="000000")
    font_sub_header = Font(name="Segoe UI", size=10, bold=True, color="000000")
    font_regular = Font(name="Segoe UI", size=10, color="000000")
    
    fill_header = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
    thin_border_side = Side(border_style="thin", color="D3D3D3")
    border_cell = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1. Header Structure
    # Top Row (Row 1): Date, Kiplings, Casa Loco Express
    ws.merge_cells("A1:A3")
    ws["A1"] = "Date"
    
    ws.merge_cells("B1:E1")
    ws["B1"] = "Kiplings"
    
    ws.merge_cells("F1:K1")
    ws["F1"] = "Casa Loco Express"

    # Platform Row (Row 2)
    ws.merge_cells("B2:C2")
    ws["B2"] = "Google"
    
    ws.merge_cells("D2:E2")
    ws["D2"] = "Zomato"
    
    ws.merge_cells("F2:G2")
    ws["F2"] = "Google"
    
    ws.merge_cells("H2:I2")
    ws["H2"] = "Zomato - Delivery"
    
    ws.merge_cells("J2:K2")
    ws["J2"] = "Swiggy - Delivery"

    # Metric Row (Row 3)
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

    # Apply Header Styles to Rows 1-3
    for r in range(1, 4):
        ws.row_dimensions[r].height = 24
        for c in range(1, 12):
            cell = ws.cell(row=r, column=c)
            cell.font = font_header_title if r < 3 else font_sub_header
            cell.alignment = align_center
            cell.border = border_cell
            cell.fill = fill_header

    # 2. Write Data Rows
    current_row = 4
    for row in rows_data:
        ws.row_dimensions[current_row].height = 20
        for col_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font = font_regular
            cell.border = border_cell
            cell.alignment = align_center
            
            # Format numbers
            if col_idx in [3, 5, 7, 9, 11] and isinstance(val, (int, float)): # Rating columns
                cell.number_format = "0.0"
            elif col_idx in [2, 4, 6, 8, 10] and isinstance(val, int): # Count columns
                cell.number_format = "#,##0"
                
        current_row += 1

    # 3. Adjust Column Widths
    ws.column_dimensions['A'].width = 14
    for col_letter in ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']:
        ws.column_dimensions[col_letter].width = 13

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
