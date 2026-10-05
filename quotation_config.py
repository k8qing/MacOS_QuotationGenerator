"""
quotation_config.py
===================
EVERYTHING you normally need to edit lives in this file (plus products.csv).
The layout/building code is in generate_quotation.py -- you should not need to touch it.

Mini text-markup used in the text fields below (toggle on/off, can be nested):
    **bold**   ~~italic~~   __underline__   ^^red^^   @@blue link look@@
Example:  "Within ^^45 -60 working days^^ upon receipt"  ->  the dates are red.
"""

import sys
from pathlib import Path

# When packaged (.exe on Windows, .app on Mac) the read-only files (logo, signature) live inside the
# app (RESOURCE_DIR) while the files you edit (products.csv, images, templates, output ...) live in APP_DIR:
#   Windows / Linux : next to the .exe
#   Mac             : ~/Documents/Keiths Quotation Generator   (a .app bundle must not be written to)
FROZEN = getattr(sys, "frozen", False)
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))
MAC_DATA_FOLDER_NAME = "Keiths Quotation Generator"


def find_app_dir(frozen=FROZEN, platform=sys.platform, executable=sys.executable, source_dir=Path(__file__).parent):
    """Folder that holds the files the user edits."""
    if not frozen:
        return Path(source_dir)
    if platform == "darwin":
        folder = Path.home() / "Documents" / MAC_DATA_FOLDER_NAME
        try:
            folder.mkdir(parents=True, exist_ok=True)
            return folder
        except OSError:
            return Path.home()
    return Path(executable).parent


APP_DIR = find_app_dir()
BASE = APP_DIR

if FROZEN:
    # python-docx opens its templates through "docx/parts/../templates/...", so a "parts" folder must exist.
    # build_exe.py bundles one; this only covers builds made without it (and fails quietly if read-only).
    try:
        (RESOURCE_DIR / "docx" / "parts").mkdir(parents=True, exist_ok=True)
    except OSError:
        pass

# ----------------------------------------------------------------------------
# 1. FILL-IN-THE-BLANK FIELDS (top of the quotation)
#    Leave "" for a blank field.  Date "" = today's date.
#    These can also be overridden from the command line (see README / --help).
# ----------------------------------------------------------------------------
QUOTE_INFO = {
    "quotation_number": "Q-26-626",
    "date":             "October 2, 2026",
    "company_name":     "",
    "attention_to":     "PLTCOL Fred Lenomta",
    "contact":          "+639985705604",
    "email":            "",
    "address":          "",
}

# ----------------------------------------------------------------------------
# 2. PRICING OPTIONS
# ----------------------------------------------------------------------------
# Default = None  -> NO "Total Discounted Price" row is shown.
# Set a number (e.g. 2320000) to show the red discounted-price row.
DISCOUNTED_PRICE = None



CURRENCY = "₱"
PACKAGE_TITLE = ""                 # default Package Title (the app starts blank; templates carry their own)
PACKAGE_BANNER_SUFFIX = " PACKAGE" # ORDER SUMMARY banner = title + this, e.g. "MATRICE 4 THERMAL PACKAGE"
                                   # ("" = show the title as typed; no banner row at all if the title is blank)

# ----------------------------------------------------------------------------
# 3. FILES
# ----------------------------------------------------------------------------
PRODUCTS_CSV = BASE / "products.csv"      # item list: names, SRP, pictures, details
IMAGES_DIR   = BASE / "images"            # product pictures named in the CSV 'images' column
LOGO_IMAGE   = RESOURCE_DIR / "assets" / "logo.png"
SIGNATURE_IMAGE = RESOURCE_DIR / "assets" / "signature.png"   # set to None for no signature
OUTPUT_DIR   = BASE / "output"
# Output file names, e.g.  Q-26-626 - 20261002 - MATRICE 4 THERMAL - PNP Mindoro.docx / .pdf
#   number  = Quotation Number            date    = Date as YYYYMMDD
#   package = Package Title               client  = Company Name, or Attention To if there is no company
# Re-order or delete parts as you like. Parts that are blank are skipped (no stray " - ").
FILENAME_PARTS = ("number", "date", "package", "client")
FILENAME_SEPARATOR = " - "

# ----------------------------------------------------------------------------
# 4. LETTERHEAD (repeats on every page) AND FOOTER
# ----------------------------------------------------------------------------
HEADER_LINES = [
    # (text, is_hyperlink_style)
    ("aero@enabled.ph", True),
    ("www.enabled.ph", True),
    ("Legal Entity: Agile Technologies Inc.", False),
    ("Globe – 0917 729 8994", False),
    ("Globe – 0917 186 6161", False),
    ("Landline: (02) 7000-5184", False),
]
FOOTER_ADDRESS = "Altitude Digital, Sparta PH, Pioneer Street, Barangay Highway Hills, Mandaluyong City"

HEADER_TEXT_COLOR = "467886"   # teal-grey letterhead text
FOOTER_BAR_COLOR  = "001A72"   # navy bar at the bottom of every page

# ----------------------------------------------------------------------------
# 5. ORDER SUMMARY -> TERMS & CONDITIONS
#    L(text, marker, at, text_at, after)
#      marker  : bullet symbol            at      : marker indent (inches)
#      text_at : text indent (inches)     after   : space below the line (points)
# ----------------------------------------------------------------------------
def L(text, marker="●", at=0.25, text_at=0.50, after=12):
    return dict(text=text, marker=marker, at=at, text_at=text_at, after=after)

TERMS_BEFORE_TIMELINE = [
    L("**Availability:** Pre - order"),
    L("**Delivery:** Pick up or Delivery; Delivery Fee care-off by the client.\n~~*FREE DELIVERY INSIDE METRO MANILA~~"),
    L("**Delivery Timeline:**", after=6),
]

# Delivery timeline table. Use None in the 2nd column to merge the cell with the one above it.
TIMELINE_HEADER = ("Location / Remarks", "Lead Time")
TIMELINE_ROWS = [
    ("Metro Manila (For Goods)",
     "Within ^^45 -60 working days^^ upon receipt of payment except for pre-ordered items"),
    ("Provincial via LBC (For Goods)", None),
    ("For Delivery & Training", "By Schedule upon receipt of payment"),
]

TERMS_AFTER_TIMELINE = [
    L("**Payment Terms**:", after=0),
    L("**~~^^50% DP and 50% Upon Fulfillment^^~~**", marker="o", at=0.50, text_at=0.77, after=0),
    L("**Accepted Methods:** Cash, Cheque Deposit, Manager’s Check, Bank Transfer",
      marker="o", at=0.50, text_at=0.77, after=14),
]

# Plain paragraph (no bullet) shown after the payment terms
NOTE = ("**NOTE:** ~~__Full payment is required **^^once equipment and training have been delivered "
        "and completed.^^** Processing or issuance of CAAP licensing/assistance does not affect "
        "payment schedule.__~~")

TERMS_FINAL = [
    L("**Price Validity:** 30 DAYS*", after=0),
    L("**~~*All prices quoted are subject to change without prior notice~~**",
      marker="", at=0.25, text_at=1.00, after=12),
    L("**Warranty:**", after=0),
    L("**1-Year Limited Warranty from DJI.**", marker="○", at=0.77, text_at=1.0, after=0),
    L("7 Days Replacement.", marker="■", at=1.25, text_at=1.5, after=0),
    L("Human error not covered", marker="■", at=1.25, text_at=1.5, after=0),
    L("**For Warranty Concerns:**", marker="○", at=0.77, text_at=1.0, after=0),
    L("Please email djiphilippines.servicecenter@phsmartfuture.com", marker="■", at=1.25, text_at=1.5, after=0),
    L("Contact #: 09602006581 - 09205788216 - 09205785698", marker="■", at=1.25, text_at=1.5, after=0),
    L("Message \"TekMage\" - Facebook page", marker="■", at=1.25, text_at=1.5, after=0),
    L("@@https://www.facebook.com/TekMagePhilippines@@", marker="■", at=1.25, text_at=1.5, after=14),
]

# ----------------------------------------------------------------------------
# 6. COMPANY / PAYMENT DETAILS BLOCK (after the terms)
# ----------------------------------------------------------------------------
COMPANY_BLOCK = [
    "**Company Name:** AGILE Technologies Inc.",
    "**TIN Number**: 774-133-717-000",
    "***Deposit to our BDO account**",
    "Beneficiary Account Name: Agile Technologies, Inc",
    "Beneficiary Bank Name: BDO Unibank, Inc.",
    "Beneficiary Account No.:  0121 3800 2991",
    "Beneficiary Bank Address: G/F, Units 3 & 4 Launch Pad Bldg. Reliance corner Sheridan St. Brgy, Mandaluyong",
    "",
    "***Shop Address:**",
    "-Altitude Digital -",
    "Altitude Digital, Sparta PH, Pioneer Street, Highway Hills, Mandaluyong from Monday to Saturday, between 8 AM to 5 PM.",
    "Pin: Sparta Philippines",
]

PREPARED_BY = [
    "Engr. Jezreille Keith D. Enguito",
    "Technical Sales Associate",
    "**AERO – ENABLE PH**",
]
