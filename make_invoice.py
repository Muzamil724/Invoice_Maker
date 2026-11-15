from invoice_core import find_problems, load_master
import argparse
from copy import copy
from datetime import date

import pandas as pd
from openpyxl import load_workbook

parser = argparse.ArgumentParser(description="Generate a CSD sales tax invoice from a code/qty list.")
parser.add_argument("input_csv", help="CSV with columns: code,qty")
parser.add_argument("--invoice", required=True, help="Invoice number, e.g. 1035")
parser.add_argument("--po", required=True, help="PO / order number from the CSD portal")
parser.add_argument("--place", required=True, help="Delivery place, e.g. 'CSD Peshawar'")
parser.add_argument("--date", default=date.today().strftime("%d-%m-%Y"),
                    help="Invoice date as DD-MM-YYYY (default: today)")
args = parser.parse_args()

master = load_master("data/Master_13-02-2025.xlsx")

# Keep only real product rows (drops the totals row at the bottom)
master = master[pd.to_numeric(master["Item Code"], errors="coerce").notna()].copy()
master["Item Code"] = master["Item Code"].astype(int)

print(master.shape)
print(master[["Item Code", "Description", "Unit PriceWithout GST", "Rate of Sale Tax"]].head())

# --- Step 2: read the order list and look each code up in master ---
inp = pd.read_csv(args.input_csv)

errors = find_problems(inp, master)
if errors:
    raise SystemExit("Fix these before generating:\n" + "\n".join(errors))

inp["qty"] = pd.to_numeric(inp["qty"]).astype(int)

lines = inp.merge(master, left_on="code", right_on="Item Code", how="left")
lines = lines[["Item Code", "Description", "qty", "Unit PriceWithout GST", "Rate of Sale Tax"]]

# --- Step 3b: fill the header and save a COPY ---

wb = load_workbook("data/CSD_CAV_GR.xlsx")
ws = wb.active

ws["A2"] = f"Invoice No: {args.invoice}"
ws["A3"] = f"Order No:  {args.po}"
ws["D3"] = f"Date: {args.date}"


# A5 has three lines; change only the place and keep the rest as it is
assert "CSD Cav Gr Lhr" in ws["A5"].value, "Template A5 changed: placeholder not found"
ws["A5"] = ws["A5"].value.replace("CSD Cav Gr Lhr", args.place)

# --- Step 3c: replace the sample rows with our lines ---

FIRST = 7           # first item row in the template
TEMPLATE_TOTAL = 30 # totals row in the template
COLS = "ABCDEFGHI"

# 1. Remember how an item row and the totals row look
item_style = {c: copy(ws[f"{c}{FIRST}"]._style) for c in COLS}
total_style = {c: copy(ws[f"{c}{TEMPLATE_TOTAL}"]._style) for c in COLS}
blank_style = copy(ws["K1"]._style)   # an unused cell = plain style

# 2. Unmerge the old totals row, then wipe the old rows
lines = lines.sort_values("Item Code").reset_index(drop=True)
ws.unmerge_cells(f"A{TEMPLATE_TOTAL}:C{TEMPLATE_TOTAL}")
for r in range(FIRST, max(ws.max_row, FIRST + len(lines)) + 1):
    for c in COLS:
        cell = ws[f"{c}{r}"]
        cell.value = None
        cell._style = copy(blank_style)

# 3. Write the item rows
for n, row in lines.iterrows():
    r = FIRST + n
    values = {
        "A": n + 1,
        "B": int(row["Item Code"]),
        "C": row["Description"],
        "D": int(row["qty"]),
        "E": float(row["Unit PriceWithout GST"]),
        "F": float(row["Rate of Sale Tax"]),
        "G": f"=E{r}*D{r}",
        "H": f"=F{r}*G{r}",
        "I": f"=H{r}+G{r}",
    }
    for c, v in values.items():
        ws[f"{c}{r}"] = v
        ws[f"{c}{r}"]._style = copy(item_style[c])
        if c == "E":
            ws[f"{c}{r}"].number_format = "0.00"

# 4. Totals row right after the last item
last = FIRST + len(lines) - 1
t = last + 1
for c in COLS:
    ws[f"{c}{t}"]._style = copy(total_style[c])
for c in "DGHI":
    ws[f"{c}{t}"] = f"=SUM({c}{FIRST}:{c}{last})"
ws.merge_cells(f"A{t}:C{t}")

out_path = f"output/invoice_{args.invoice}.xlsx"
wb.save(out_path)
print(f"Saved {out_path} ({len(lines)} items)")
