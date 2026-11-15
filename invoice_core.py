import pandas as pd


def load_master(path):
    """Read the master file and keep only real product rows."""
    master = pd.read_excel(path, sheet_name="Sheet1")
    master = master[pd.to_numeric(master["Item Code"], errors="coerce").notna()].copy()
    master["Item Code"] = master["Item Code"].astype(int)
    return master


def find_problems(inp, master):
    """Return a list of error messages for an order list (empty list = all good)."""
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

    return errors