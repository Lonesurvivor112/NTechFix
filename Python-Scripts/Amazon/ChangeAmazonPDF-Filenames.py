"""
Amazon Order PDF Renamer
========================
Extracts order info from Amazon order PDFs and saves renamed copies using a
clean, consistent format: MM-DD-YYYY-Amazon-genericname-OrderNumber.pdf

Folder structure:
    AmazonPDF-Tools/
    ├── ChangeAmazonPDF-Filenames.py   (this script)
    ├── SourcePDFs/               (drop your Amazon PDFs here)
    └── DestinationPDFs/          (renamed copies land here)

Requirements:
    pip install pdfplumber

Usage:
    1. Drop your Amazon order PDFs into the SourcePDFs folder.
    2. Run:  python amazon_order_renamer.py
    3. For each PDF, review the extracted info and type a generic name.
    4. The renamed copy is saved to DestinationPDFs.
       (Originals in SourcePDFs are untouched.)

Author: Built with Claude
"""

import os
import re
import sys
import glob
import shutil
import pdfplumber
from datetime import datetime


# ──────────────────────────────────────────────────────────────────────
#  CONFIGURATION — Change these to fit your setup
# ──────────────────────────────────────────────────────────────────────

# Base directory where this tool lives.
# Default: the same folder the script is in.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Source folder — drop your Amazon order PDFs here.
SOURCE_FOLDER = os.path.join(BASE_DIR, "SourcePDFs")

# Destination folder — renamed PDFs are copied here (originals stay in Source).
DESTINATION_FOLDER = os.path.join(BASE_DIR, "DestinationPDFs")

# Set to True to process ALL PDFs in the folder automatically.
# Set to False to be prompted before processing each file.
AUTO_PROCESS_ALL = False


# ──────────────────────────────────────────────────────────────────────
#  EXTRACTION FUNCTIONS
# ──────────────────────────────────────────────────────────────────────

def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract all text from a PDF file using pdfplumber."""
    full_text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    full_text += page_text + "\n"
    except Exception as e:
        print(f"  [ERROR] Could not read PDF: {e}")
    return full_text


def extract_order_number(text: str) -> str | None:
    """
    Find the Amazon order number in the extracted text.
    Amazon order numbers look like: 111-7686294-7061841
    """
    # Pattern: 3 digits, dash, 7 digits, dash, 7 digits
    match = re.search(r"\b(\d{3}-\d{7}-\d{7})\b", text)
    if match:
        return match.group(1)
    return None


def extract_order_date(text: str) -> str | None:
    """
    Find the order date and return it in MM-DD-YYYY format.
    Amazon PDFs typically show dates like:
        "Order Placed: March 14, 2026"
        "Order placed March 14, 2026"
        "Ordered on March 14, 2026"
        "March 14, 2026"  (near "Order" context)
    """
    # Try several patterns Amazon uses
    date_patterns = [
        # "Order Placed: March 14, 2026" or "Order placed March 14, 2026"
        r"[Oo]rder\s+[Pp]laced[:\s]+([A-Z][a-z]+ \d{1,2},?\s*\d{4})",
        # "Ordered on March 14, 2026"
        r"[Oo]rdered\s+on[:\s]+([A-Z][a-z]+ \d{1,2},?\s*\d{4})",
        # "Order date: March 14, 2026"
        r"[Oo]rder\s+[Dd]ate[:\s]+([A-Z][a-z]+ \d{1,2},?\s*\d{4})",
        # Fallback: any "Month Day, Year" near the word "order"
        r"[Oo]rder.*?([A-Z][a-z]+ \d{1,2},?\s*\d{4})",
        # Last resort: first occurrence of "Month Day, Year" anywhere
        r"([A-Z][a-z]+ \d{1,2},?\s*\d{4})",
    ]

    for pattern in date_patterns:
        match = re.search(pattern, text)
        if match:
            raw_date = match.group(1).strip()
            # Normalize: ensure comma is present
            raw_date = re.sub(r"(\d{1,2})\s+(\d{4})", r"\1, \2", raw_date)
            try:
                parsed = datetime.strptime(raw_date, "%B %d, %Y")
                return parsed.strftime("%m-%d-%Y")
            except ValueError:
                continue

    # Try numeric date formats as well (e.g., 03/14/2026)
    numeric_patterns = [
        r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})",
    ]
    for pattern in numeric_patterns:
        match = re.search(pattern, text)
        if match:
            try:
                month, day, year = match.group(1), match.group(2), match.group(3)
                parsed = datetime(int(year), int(month), int(day))
                return parsed.strftime("%m-%d-%Y")
            except ValueError:
                continue

    return None


def extract_po_number(text: str) -> str | None:
    """
    Extract the PO number from Amazon Business order PDFs.
    These appear as "PO number : some text" in the header.
    """
    match = re.search(r"PO\s+number\s*:\s*(.+?)(?:\n|$)", text, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    return None


def extract_items(text: str) -> list[str]:
    """
    Extract item titles from Amazon order PDF text.
    This uses heuristics since Amazon PDF formats can vary.
    Returns a list of strings like "Item Name" or "(x2) Item Name".
    """
    items = []
    lines = text.split("\n")

    # Strategy 1: Lines that appear after quantity indicators
    # Amazon PDFs often have lines like "1 of: Item Name" or "Qty: 1  Item Name"
    for i, line in enumerate(lines):
        line_stripped = line.strip()

        # Pattern: "2 of: Product Name" or "Qty: 2 Product Name"
        qty_match = re.match(r"(?:(\d+)\s+of[:\s]+|Qty[:\s]*(\d+)\s+)(.+)", line_stripped, re.IGNORECASE)
        if qty_match:
            qty = qty_match.group(1) or qty_match.group(2)
            item_name = qty_match.group(3).strip()
            # Strip trailing price (e.g. "$69.99" at end of line)
            item_name = re.sub(r"\s*\$[\d,]+\.\d{2}\s*$", "", item_name).strip()
            if len(item_name) > 5:  # Skip very short strings (unlikely item names)
                if qty and int(qty) > 1:
                    items.append(f"(x{qty}) {item_name}")
                else:
                    items.append(item_name)

    # Strategy 2: Look for lines near "Sold by" or "Shipped" markers
    if not items:
        for i, line in enumerate(lines):
            line_stripped = line.strip()
            # If we see "Sold by" or "Condition:", the previous non-empty line
            # is likely the item name
            if re.match(r"(Sold by|Condition:|Fulfilled by)", line_stripped, re.IGNORECASE):
                # Look backwards for the item name
                for j in range(i - 1, max(i - 4, -1), -1):
                    candidate = lines[j].strip()
                    if (
                        len(candidate) > 10
                        and not re.match(r"(Order|Shipping|Payment|Delivery|Sold by|Condition)", candidate, re.IGNORECASE)
                        and not re.match(r"^\$", candidate)
                    ):
                        if candidate not in items:
                            items.append(candidate)
                        break

    # Strategy 3: Broad fallback — look for long descriptive lines
    # that appear between the order number and the totals section
    if not items:
        in_items_section = False
        for line in lines:
            line_stripped = line.strip()
            # Start capturing after order number
            if re.search(r"\d{3}-\d{7}-\d{7}", line_stripped):
                in_items_section = True
                continue
            # Stop at totals/payment section
            if re.match(r"(Item.s.? Subtotal|Order Total|Payment|Grand Total|Subtotal)", line_stripped, re.IGNORECASE):
                in_items_section = False
                continue
            if in_items_section and len(line_stripped) > 20:
                # Skip lines that look like prices, addresses, or dates
                if not re.match(r"^[\$\d]", line_stripped) and \
                   not re.match(r"(Shipping|Delivery|Payment|Sold|Condition|Guarantee)", line_stripped, re.IGNORECASE):
                    items.append(line_stripped)

    return items


# ──────────────────────────────────────────────────────────────────────
#  RENAMING LOGIC
# ──────────────────────────────────────────────────────────────────────

def validate_generic_name(name: str) -> str | None:
    """
    Validate the user-provided generic name.
    Rules: lowercase, no spaces, no special chars, not empty.
    Returns cleaned name or None if invalid.
    """
    name = name.strip().lower()
    if not name:
        return None
    # Remove anything that isn't a letter, digit, or hyphen
    cleaned = re.sub(r"[^a-z0-9\-]", "", name)
    if not cleaned:
        return None
    return cleaned


def build_new_filename(date_str: str, generic_name: str, order_number: str) -> str:
    """Build the standardized filename."""
    return f"{date_str}-Amazon-{generic_name}-{order_number}.pdf"


# ──────────────────────────────────────────────────────────────────────
#  MAIN WORKFLOW
# ──────────────────────────────────────────────────────────────────────

def process_single_pdf(pdf_path: str) -> bool:
    """
    Process one Amazon order PDF:
    1. Extract order info
    2. Display it to the user
    3. Ask for generic name
    4. Copy the renamed file to DestinationPDFs

    Returns True if processed, False if skipped.
    """
    filename = os.path.basename(pdf_path)

    print(f"\n{'='*60}")
    print(f"  FILE: {filename}")
    print(f"{'='*60}")

    # Extract text
    text = extract_text_from_pdf(pdf_path)
    if not text.strip():
        print("  [SKIP] Could not extract any text from this PDF.")
        print("         It may be a scanned image or not an Amazon order.")
        return False

    # Extract order details
    order_number = extract_order_number(text)
    order_date = extract_order_date(text)
    po_number = extract_po_number(text)
    items = extract_items(text)

    # Display extracted info
    print()
    if order_number:
        print(f"  Order Number : {order_number}")
    else:
        print("  Order Number : [NOT FOUND]")

    if order_date:
        print(f"  Order Date   : {order_date}")
    else:
        print("  Order Date   : [NOT FOUND]")

    if po_number:
        print(f"  PO Number    : {po_number}")

    if items:
        print(f"\n  Items found ({len(items)}):")
        for item in items:
            print(f"    - {item}")
    else:
        print("\n  Items: [NONE DETECTED — check PDF manually]")

    print()

    # If we're missing critical info, let user provide it manually
    if not order_number:
        print("  Could not find the order number automatically.")
        user_input = input("  Enter order number (or 'skip' to skip this file): ").strip()
        if user_input.lower() == "skip":
            print("  [SKIPPED]")
            return False
        order_number = user_input

    if not order_date:
        print("  Could not find the order date automatically.")
        user_input = input("  Enter date as MM-DD-YYYY (or 'skip'): ").strip()
        if user_input.lower() == "skip":
            print("  [SKIPPED]")
            return False
        # Validate format
        try:
            parsed = datetime.strptime(user_input, "%m-%d-%Y")
            order_date = parsed.strftime("%m-%d-%Y")
        except ValueError:
            print("  [ERROR] Invalid date format. Skipping this file.")
            return False

    # Ask for generic name
    while True:
        generic_name = input("  Enter generic name for this order: ").strip()
        validated = validate_generic_name(generic_name)
        if validated:
            break
        print("  [INVALID] Use lowercase letters/numbers only, no spaces.")
        print("  Examples: monitor, cable, dock, usb-hub")

    # Build new filename and destination path
    new_filename = build_new_filename(order_date, validated, order_number)
    new_path = os.path.join(DESTINATION_FOLDER, new_filename)

    # Check for conflicts
    if os.path.exists(new_path):
        print(f"\n  [WARNING] File already exists in DestinationPDFs: {new_filename}")
        overwrite = input("  Overwrite? (y/n): ").strip().lower()
        if overwrite != "y":
            print("  [SKIPPED]")
            return False

    # Copy renamed file to destination
    try:
        shutil.copy2(pdf_path, new_path)
        print(f"\n  [DONE] Saved to DestinationPDFs:")
        print(f"         {new_filename}")
        return True
    except OSError as e:
        print(f"\n  [ERROR] Could not rename file: {e}")
        return False


def main():
    """Main entry point — find PDFs and process them."""
    print()
    print("=" * 60)
    print("  AMAZON ORDER PDF RENAMER")
    print("=" * 60)
    print(f"  Base dir     : {BASE_DIR}")
    print(f"  Source       : {SOURCE_FOLDER}")
    print(f"  Destination  : {DESTINATION_FOLDER}")
    print()

    # Create folders if they don't exist
    os.makedirs(SOURCE_FOLDER, exist_ok=True)
    os.makedirs(DESTINATION_FOLDER, exist_ok=True)

    # Find all PDFs in the source folder
    search_path = os.path.join(SOURCE_FOLDER, "*.pdf")
    pdf_files = sorted(glob.glob(search_path))

    if not pdf_files:
        print("  No PDF files found in SourcePDFs.")
        print(f"  Drop your Amazon order PDFs into:")
        print(f"    {SOURCE_FOLDER}")
        sys.exit(0)

    # Filter out already-renamed files (match our naming pattern)
    already_renamed_pattern = re.compile(r"^\d{2}-\d{2}-\d{4}-Amazon-.+\.pdf$", re.IGNORECASE)
    unprocessed = []
    skipped_count = 0

    for f in pdf_files:
        basename = os.path.basename(f)
        if already_renamed_pattern.match(basename):
            skipped_count += 1
        else:
            unprocessed.append(f)

    print(f"  Found {len(pdf_files)} PDF(s) in SourcePDFs")
    if skipped_count > 0:
        print(f"  Skipping {skipped_count} already-renamed file(s)")
    print(f"  Processing {len(unprocessed)} file(s)")

    if not unprocessed:
        print("\n  Nothing to process — all files are already renamed!")
        sys.exit(0)

    # Process each PDF
    renamed_count = 0
    for i, pdf_path in enumerate(unprocessed, 1):
        if not AUTO_PROCESS_ALL and i > 1:
            cont = input(f"\n  Continue to next file? ({i}/{len(unprocessed)}) [y/n]: ").strip().lower()
            if cont != "y":
                print("  Stopping.")
                break

        if process_single_pdf(pdf_path):
            renamed_count += 1

    # Summary
    print(f"\n{'='*60}")
    print(f"  DONE — Processed {renamed_count} of {len(unprocessed)} file(s)")
    print(f"  Renamed files saved to: {DESTINATION_FOLDER}")
    print(f"{'='*60}")
    print()


if __name__ == "__main__":
    main()
