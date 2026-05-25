import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side,
    GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import datetime

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Apartment Tracker"

# ── Colour palette ──────────────────────────────────────────────────────────
NAVY        = "1F3864"   # header row background
GOLD        = "C9A84C"   # header text
LIGHT_BLUE  = "DCE6F1"   # even row shading
WHITE       = "FFFFFF"
GREEN_FILL  = "E2EFDA"   # purchase rows
YELLOW_FILL = "FFF2CC"   # rental rows
ORANGE_FILL = "FCE4D6"   # passed rows
BLUE_FILL   = "DDEBF7"   # contacted rows (awaiting / scheduling)

# ── Column definitions ────────────────────────────────────────────────────
columns = [
    ("Address",               28),
    ("Unit",                  10),
    ("StreetEasy Link",       35),
    ("Type",                  12),   # Rent / Purchase
    ("Asking Price",          15),
    ("Offer Price",           15),
    ("Broker Name",           24),
    ("Broker Email",          32),
    ("Date Contacted",        16),
    ("Date of Showing",       22),   # date + time
    ("Status",                24),
    ("Notes",                 55),
]

# ── Header row ───────────────────────────────────────────────────────────
header_font    = Font(name="Calibri", bold=True, color=GOLD, size=11)
header_fill    = PatternFill("solid", fgColor=NAVY)
header_align   = Alignment(horizontal="center", vertical="center", wrap_text=True)
thin           = Side(style="thin", color="AAAAAA")
thin_border    = Border(left=thin, right=thin, top=thin, bottom=thin)

ws.row_dimensions[1].height = 32
for col_idx, (col_name, col_width) in enumerate(columns, start=1):
    cell = ws.cell(row=1, column=col_idx, value=col_name)
    cell.font      = header_font
    cell.fill      = header_fill
    cell.alignment = header_align
    cell.border    = thin_border
    ws.column_dimensions[get_column_letter(col_idx)].width = col_width

# ── Freeze the header row + add auto-filter ──────────────────────────────
ws.freeze_panes = "A2"
ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}1"

# ── Apartment data ───────────────────────────────────────────────────────
# Fields: address, unit, se_link, type, asking, offer,
#         broker, email, date_contacted, showing_date, status, notes
apartments = [
    # ── PURCHASE ─────────────────────────────────────────────────────────
    (
        "200 East 90th Street, New York, NY 10128",
        "16E",
        "https://streeteasy.com/building/200-east-90th-street-new_york",
        "Purchase",
        "$1,150,000",
        "$1,020,000",
        "Sean Devine (Devine Team at Compass)",
        "sean.devine@compass.com",
        "Jan 2026",
        "May 3, 2026 at 11:00 AM",
        "Active – Negotiating",
        (
            "Co-op. Asking reduced from $1.25M → $1.15M. "
            "Offer $1,020,000; seller countered at $1,125,000. "
            "High floor, East River views. "
            "Buyer's broker: Saree Ptak (sareerptak@gmail.com)."
        ),
    ),
    (
        "10 East End Avenue, New York, NY 10028",
        "18A",
        "https://streeteasy.com/building/10-east-end-avenue-new_york",
        "Purchase",
        "$1,150,000",
        "$1,150,000",
        "Jonathan Tauzowicz (The Collective Team at Compass)",
        "jt@compass.com",
        "Jan 2026",
        "Jan 12, 2026 at 9:45 AM",
        "Under Contract",
        (
            "Co-op. Estate sale (Estate of Tuppatsch). "
            "Agreed price $1,150,000. Contract in progress with attorney Andy Cutler. "
            "Court action (estate) adjourned to 5/5/26. "
            "Chase mortgage (7yr ARM 4.750%). "
            "Buyer's broker: Saree Ptak (sareerptak@gmail.com)."
        ),
    ),
    (
        "520 East 72nd Street, New York, NY 10021",
        "PHC",
        "https://streeteasy.com/building/520-east-72nd-street-new_york",
        "Purchase",
        "",
        "",
        "Louise Devlin (Brown Harris Stevens)",
        "LDevlin@bhsusa.com",
        "Jan 2026",
        "Jan 11, 2026",
        "Inquired",
        (
            "Co-op. 30% down building; no W/D in unit; no investors allowed. "
            "DTI < 30% required. 2% buyer's agent commission. "
            "Buyer's broker: Saree Ptak (sareerptak@gmail.com)."
        ),
    ),
    # ── RENTAL ───────────────────────────────────────────────────────────
    (
        "265 East 66th Street, New York, NY 10065",
        "34B",
        "https://streeteasy.com/building/265-east-66th-street-new_york",
        "Rent",
        "",
        "",
        "Joseph George (MNS Real Estate)",
        "jgj@mns.com",
        "May 25, 2026",
        "May 31, 2026 at 12:00 PM",
        "Showing Scheduled",
        (
            "Rental inquiry sent May 25, 2026 (via Harley's StreetEasy account). "
            "Showing confirmed: Sunday May 31, 2026 at 12:00 PM ET. "
            "Contact: Joseph George, MNS Real Estate (jgj@mns.com)."
        ),
    ),
    (
        "240 East 82nd Street, New York, NY 10028",
        "18B",
        "https://streeteasy.com/building/240-east-82nd-street-new_york",
        "Rent",
        "$7,000/mo",
        "",
        "Tomi Bicanic (Citywide Apts)",
        "tomi.bicanic@citywideapts.com",
        "May 24, 2026",
        "TBD – Sat May 30/31 requested",
        "Contacted – Scheduling",
        (
            "StreetEasy inquiry sent May 24, 2026. $7,000/mo. "
            "Celine requested Saturday tour (out of town until Sat). "
            "Contact: Tomi Bicanic, Citywide Apts (tomi.bicanic@citywideapts.com)."
        ),
    ),
    (
        "345 East 80th Street, New York, NY 10075",
        "27B",
        "https://streeteasy.com/building/345-east-80th-street-new_york",
        "Rent",
        "",
        "",
        "Anis (PocketBroker)",
        "anis@pocketbroker.com",
        "May 24, 2026",
        "TBD",
        "Contacted – Circle Back When Ready",
        (
            "StreetEasy inquiry sent May 24, 2026. "
            'Broker replied: "Enjoy your travels. Circle back to us then." '
            "Also cc: diana@pocketbroker.com. "
            "Follow up when back in town."
        ),
    ),
    (
        "170 East 87th Street, New York, NY 10128",
        "W4H",
        "https://streeteasy.com/building/170-east-87th-street-new_york",
        "Rent",
        "",
        "",
        "",
        "",
        "May 24, 2026",
        "TBD",
        "Contacted – Awaiting Response",
        (
            "StreetEasy inquiry sent May 24, 2026. "
            "No broker reply received as of May 25, 2026. "
            "Broker name/email TBD."
        ),
    ),
]

# ── Row styles keyed by type / status ────────────────────────────────────
def row_fill(apt_type, status):
    s = status.lower()
    if "passed" in s:
        return PatternFill("solid", fgColor=ORANGE_FILL)
    if "contacted" in s or "scheduling" in s:
        return PatternFill("solid", fgColor=BLUE_FILL)
    if apt_type.lower() == "purchase":
        return PatternFill("solid", fgColor=GREEN_FILL)
    return PatternFill("solid", fgColor=YELLOW_FILL)

data_font     = Font(name="Calibri", size=10)
data_align_ww = Alignment(vertical="top", wrap_text=True)
data_align_c  = Alignment(horizontal="center", vertical="top")
link_font     = Font(name="Calibri", size=10, color="0563C1", underline="single")

for row_idx, apt in enumerate(apartments, start=2):
    (address, unit, se_link, apt_type, asking, offer,
     broker, email, date_contacted, showing, status, notes) = apt

    fill = row_fill(apt_type, status)
    ws.row_dimensions[row_idx].height = 65

    values = [address, unit, se_link, apt_type, asking, offer,
              broker, email, date_contacted, showing, status, notes]

    for col_idx, value in enumerate(values, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.font   = data_font
        cell.fill   = fill
        cell.border = thin_border

        if col_idx == 3 and value:          # StreetEasy link – hyperlink
            cell.font      = link_font
            cell.hyperlink = value
            cell.value     = value
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        elif col_idx in (1, 7, 12):         # wrap text: Address, Broker Name, Notes
            cell.alignment = data_align_ww
        elif col_idx in (2, 4, 5, 6, 9, 10, 11):  # centered: Unit, Type, Prices, Dates, Status
            cell.alignment = data_align_c
        else:
            cell.alignment = data_align_ww

# ── Instructions / Legend tab ─────────────────────────────────────────────
ws2 = wb.create_sheet("How to Use")
ws2.column_dimensions["A"].width = 85

instructions = [
    ("NYC APARTMENT TRACKER – HOW TO USE", True),
    ("", False),
    ("EDITING FROM EXCEL", True),
    ("• Open NYC_Apartment_Tracker.xlsx directly and edit any cell.", False),
    ("• To add a new apartment: copy any existing row and paste at the bottom, then fill in the fields.", False),
    ("• Save when done. The file is stored in the repo.", False),
    ("", False),
    ("EDITING FROM CLAUDE", True),
    ('  Example: "Add 123 Main St, Unit 4A, asking $1.2M, purchase, broker Jane Smith at jane@compass.com, showing June 1 at 2pm"', False),
    ('  Example: "Update the notes for 10 East End Ave 18A"', False),
    ('  Example: "Mark 345 East 80th #27B as Passed"', False),
    ("• Claude reads this file, applies your change, saves, and commits.", False),
    ("", False),
    ("UPLOADING A SCREENSHOT OR EMAIL TO CHECK IF IT'S IN THE LIST", True),
    ("• Paste or drag an image / screenshot into the Claude chat.", False),
    ('  Ask: "Is this apartment already in my tracker?"', False),
    ("• Claude will read the address from the image and search the spreadsheet for a match.", False),
    ("", False),
    ("STATUS KEY", True),
    ("• Active – Negotiating         →  offer submitted, back-and-forth in progress", False),
    ("• Under Contract               →  agreed price, attorney review / board package stage", False),
    ("• Showing Scheduled            →  showing confirmed with date & time", False),
    ("• Contacted – Scheduling       →  broker replied, arranging showing date", False),
    ("• Contacted – Awaiting Response →  message sent, no reply yet", False),
    ('• Contacted – Circle Back When Ready → broker said to reach out later', False),
    ("• Inquired                     →  reached out, waiting or toured only", False),
    ("• Passed                       →  decided not to proceed", False),
    ("", False),
    ("ROW COLOUR KEY", True),
    ("• Green  →  Purchase listing, active or under contract", False),
    ("• Yellow →  Rental listing, showing scheduled or inquired", False),
    ("• Blue   →  Contacted (rental, awaiting reply or scheduling)", False),
    ("• Orange →  Passed / no longer pursuing", False),
]

title_font = Font(name="Calibri", bold=True, size=12)
body_font  = Font(name="Calibri", size=10)

for r_idx, (text, is_title) in enumerate(instructions, start=1):
    cell = ws2.cell(row=r_idx, column=1, value=text)
    cell.font      = title_font if is_title else body_font
    cell.alignment = Alignment(wrap_text=True)

output_path = "/home/user/Default/NYC_Apartment_Tracker.xlsx"
wb.save(output_path)
print(f"Saved → {output_path}")
print(f"Apartments logged: {len(apartments)}")
