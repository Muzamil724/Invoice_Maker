import pandas as pd

master = pd.read_excel("data/Master_13-02-2025.xlsx", sheet_name="Sheet1")

# Keep only real product rows (drops the totals row at the bottom)
master = master[pd.to_numeric(master["Item Code"], errors="coerce").notna()].copy()
master["Item Code"] = master["Item Code"].astype(int)

print(master.shape)
print(master[["Item Code", "Description", "Unit PriceWithout GST", "Rate of Sale Tax"]].head())

# --- Step 2: read the order list and look each code up in master ---
inp = pd.read_csv("input/test.csv")
errors = []

# Check 1: qty must be a whole number greater than 0
qty_num = pd.to_numeric(inp["qty"], errors="coerce")
bad_qty = inp[qty_num.isna() | (qty_num <= 0) | (qty_num % 1 != 0)]
for i, row in bad_qty.iterrows():
    errors.append(f"Line {i + 2}: invalid qty {row['qty']!r} (code {row['code']})")

# Check 2: every code must exist in master
unknown = inp[~inp["code"].isin(master["Item Code"])]
for i, row in unknown.iterrows():
    errors.append(f"Line {i + 2}: code {row['code']} not in master")

# Check 3: the same code must not appear twice
dupes = inp[inp["code"].duplicated(keep=False)]
for i, row in dupes.iterrows():
    errors.append(f"Line {i + 2}: code {row['code']} appears more than once")

# One exit point, after all checks have run
if errors:
    raise SystemExit("Fix these before generating:\n" + "\n".join(errors))

inp["qty"] = qty_num.astype(int)

lines = inp.merge(master, left_on="code", right_on="Item Code", how="left")
lines = lines[["Item Code", "Description", "qty", "Unit PriceWithout GST", "Rate of Sale Tax"]]
print(lines)

# --- Step 3a: inspect the invoice template ---
from openpyxl import load_workbook

wb = load_workbook("data/CSD_CAV_GR.xlsx")
ws = wb.active

print("Sheet:", ws.title, "| used range:", ws.dimensions)
print("Merged ranges:", sorted(str(r) for r in ws.merged_cells.ranges))

for addr in ["A2", "A3", "D3", "A5"]:
    print(addr, "=", repr(ws[addr].value))

for addr in ["A6", "A7", "D7", "E7", "F7", "G7", "H7", "I7", "D30", "I30"]:
    c = ws[addr]
    print(addr, "| value:", c.value, "| format:", c.number_format, "| bold:", c.font.b)


# --- Step 3b: fill the header and save a COPY ---
from datetime import date
from openpyxl import load_workbook

invoice_no = "TEST-001"
po_no = "12345678912345"
place = "CSD Peshawar"
inv_date = date.today().strftime("%d-%m-%Y")

wb = load_workbook("data/CSD_CAV_GR.xlsx")
ws = wb.active

ws["A2"] = f"Invoice No: {invoice_no}"
ws["A3"] = f"Order No:  {po_no}"
ws["D3"] = f"Date: {inv_date}"

# A5 has three lines; change only the place and keep the rest as it is
assert "CSD Cav Gr Lhr" in ws["A5"].value, "Template A5 changed: placeholder not found"
ws["A5"] = ws["A5"].value.replace("CSD Cav Gr Lhr", place)

# --- Step 3c: replace the sample rows with our lines ---
from copy import copy

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

# 4. Totals row right after the last item
last = FIRST + len(lines) - 1
t = last + 1
for c in COLS:
    ws[f"{c}{t}"]._style = copy(total_style[c])
for c in "DGHI":
    ws[f"{c}{t}"] = f"=SUM({c}{FIRST}:{c}{last})"
ws.merge_cells(f"A{t}:C{t}")

wb.save("output/test_invoice.xlsx")
print("saved")
