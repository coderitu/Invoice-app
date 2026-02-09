from datetime import datetime
from fpdf import FPDF
import os
import re
from num2words import num2words
import sys

def resource_path(relative_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.abspath(relative_path)


# DEFAULT_LOGO_PATH = os.path.abspath("Wadi_logo.png")
DEFAULT_LOGO_PATH = resource_path("Wadi_logo.png")
STEAM_SAUNA_FEE = 3000.00

# ---------------- HELPERS ----------------
def sanitize_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", str(name))


def safe_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def amount_in_words(amount):
    words = num2words(amount, lang="en")
    return f"{words.replace('-', ' ').title()} Kenya Shillings"


# ---------------- PDF GENERATOR ----------------
def generate_invoice_pdf(
    member,
    name,
    invoice_no,
    renewal,
    pending_2025,
    pending_2024,
    pending_2023,
    discount_percent=0,
    output_folder="invoices",
    logo_path=DEFAULT_LOGO_PATH
):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    filepath = os.path.abspath(
        os.path.join(
            output_folder,
            f"{sanitize_filename(invoice_no)}_{sanitize_filename(name)}.pdf"
        )
    )

    pdf = FPDF()
    pdf.add_page()

    # ---------- LOGO ----------
    if os.path.exists(logo_path):
        logo_w = 120
        pdf.image(logo_path, x=(210 - logo_w) / 2, y=8, w=logo_w)
        pdf.ln(40)

    # ---------- HEADER ----------
    invoice_date = datetime.today().strftime("%d %B %Y")

    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 8, "INVOICE", ln=True, align="R")

    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, f"Invoice No: {invoice_no}", ln=True, align="R")
    pdf.cell(0, 6, f"Date: {invoice_date}", ln=True, align="R")
    pdf.cell(0, 6, "PIN No.: P051518511A", ln=True, align="R")
    pdf.ln(8)

    # ---------- CLIENT ----------
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 6, "Invoice To:", ln=True)

    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 6, name, ln=True)
    pdf.cell(0, 6, f"Member Code: {member}", ln=True)
    pdf.ln(6)

    # ---------- TABLE DATA ----------
    rows = []
    subtotal = 0.0

    renewal_val = safe_float(renewal)
    if renewal_val > 0:
        rows.append(("Annual Renewal", renewal_val))
        subtotal += renewal_val

    rows.append(("Steam & Sauna", STEAM_SAUNA_FEE))
    subtotal += STEAM_SAUNA_FEE

    for label, value in [
        ("Pending 2025", pending_2025),
        ("Pending 2024", pending_2024),
        ("Pending 2023", pending_2023),
    ]:
        val = safe_float(value)
        if val > 0:
            rows.append((label, val))
            subtotal += val

    discount_percent = safe_float(discount_percent)
    discount_amount = subtotal * (discount_percent / 100) if discount_percent > 0 else 0
    total = subtotal - discount_amount

    # ---------- TABLE ----------
    DESC_W = 140
    AMT_W = 50
    ROW_H = 8

    pdf.set_font("Arial", "B", 10)
    pdf.cell(DESC_W, ROW_H, "Description", 1)
    pdf.cell(AMT_W, ROW_H, "Amount (KES)", 1, ln=True)

    pdf.set_font("Arial", "", 10)
    for label, value in rows:
        pdf.cell(DESC_W, ROW_H, label, 1)
        pdf.cell(AMT_W, ROW_H, f"{value:,.2f}", 1, ln=True)

    # Discount row (ONLY if used)
    if discount_percent > 0:
        pdf.cell(DESC_W, ROW_H, f"Discount ({discount_percent:.0f}%)", 1)
        pdf.cell(AMT_W, ROW_H, f"-{discount_amount:,.2f}", 1, ln=True)

    # Total
    pdf.set_font("Arial", "B", 10)
    pdf.cell(DESC_W, ROW_H, "Total", 1)
    pdf.cell(AMT_W, ROW_H, f"{total:,.2f}", 1, ln=True)

    pdf.ln(6)

    # ---------- AMOUNT IN WORDS ----------
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 6, "Amount in Words:", ln=True)

    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(0, 6, f"{amount_in_words(int(total))} Only")

    pdf.ln(8)

    # ---------- BANK DETAILS ----------
    box_x = 10
    box_y = pdf.get_y()
    box_w = 90
    box_h = 70

    pdf.rect(box_x, box_y, box_w, box_h)

    pdf.set_xy(box_x + 2, box_y + 2)
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 6, "Bank Details", ln=True)

    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(
        box_w - 4, 6,
        "Wadi Degla Investment Kenya Ltd.\n"
        "Commercial International Bank (CIB)\n"
        "Branch: Mayfair\n"
        "Branch Code: 065-001\n"
        "A/C Number: 0101350368"
    )

    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 6, "MPESA", ln=True)

    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(box_w - 4, 6, f"Paybill: 850 107\nA/C No.: {member}")

    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 6, "Cheque Details", ln=True)

    pdf.set_font("Arial", "", 10)
    pdf.multi_cell(box_w - 4, 6, "Wadi Degla Investment K Limited")

    # ---------- COMMENTS ----------
    pdf.set_xy(110, box_y)
    pdf.set_font("Arial", "B", 10)
    pdf.cell(0, 6, "Comments:", ln=True)

    pdf.set_font("Arial", "", 10)
    pdf.set_x(110)
    pdf.multi_cell(
        90, 6,
        "Indicate the invoice number in bank transfers.\n\n"
        "Client Acceptance:"
    )

    # ---------- FOOTER ----------
    divider_y = pdf.h - 40
    pdf.set_line_width(0.5)
    pdf.line(10, divider_y, pdf.w - 10, divider_y)

    pdf.set_y(divider_y + 3)
    pdf.set_font("Arial", "", 9)
    pdf.cell(
        0, 5,
        "Tel +254 792 888 888   Wadi Degla Investment (K) Ltd.   www.wadidegla.com",
        ln=True,
        align="C"
    )
    pdf.cell(0, 5, "Part of Wadi Degla Holding", ln=True, align="C")
    pdf.cell(0, 5, "41973-00100 Nairobi, Kenya", ln=True, align="C")

    pdf.output(filepath)
    return filepath
