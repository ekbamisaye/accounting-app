import sqlite3
from pathlib import Path
from datetime import date
import calendar
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "accounting_records.db"
SUMMARY_TEMPLATE = APP_DIR / "ACCOUNT SUMMARIES AS AT SEPTEMBER 2026.xlsx"
WELFARE_TEMPLATE = APP_DIR / "WELAFRE SUMMARY.xlsx"
MONTHS = list(calendar.month_name)[1:]

st.set_page_config(page_title="Account Summaries", page_icon="📒", layout="wide")

def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with connect() as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                txn_date TEXT NOT NULL,
                txn_type TEXT NOT NULL,
                category TEXT NOT NULL,
                amount REAL NOT NULL CHECK(amount >= 0),
                narration TEXT NOT NULL DEFAULT '',
                detail_name TEXT NOT NULL DEFAULT '',
                purpose TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
        """)
        # Keep any test transactions already entered under the prototype's old income labels.
        legacy_income_labels = {
            "PROPHET OFFERING": "Prophet Offering (Cash & Cheques)",
            "BANK": "Bank transfers",
            "RENT": "Rent",
            "OTHERS (DIVIDENDS, ETC.)": "Other dividends",
            "FOREX": "Forex",
        }
        for old_label, new_label in legacy_income_labels.items():
            con.execute(
                "UPDATE transactions SET category=? WHERE txn_type='INCOME' AND category=?",
                (new_label, old_label),
            )

def read_categories():
    # Friendly entry labels requested by the user; export/summary maps these to the workbook labels.
    income = ["Prophet Offering (Cash & Cheques)", "Bank transfers", "Rent", "Other dividends", "Forex"]
    expenses = []
    budgets = {}
    if SUMMARY_TEMPLATE.exists():
        wb = load_workbook(SUMMARY_TEMPLATE, data_only=True, read_only=True)
        if "INCOME" in wb.sheetnames:
            ws = wb["INCOME"]
            # Keep the requested user-facing labels rather than raw template headings.
            income = ["Prophet Offering (Cash & Cheques)", "Bank transfers", "Rent", "Other dividends", "Forex"]
        if "EXPENDITURE" in wb.sheetnames:
            ws = wb["EXPENDITURE"]
            for r in list(range(3, 47)) + list(range(51, 95)):
                desc = ws.cell(r, 3).value
                if desc and str(desc).strip():
                    expenses.append(str(desc).strip())
        if "BUDGET (INCOME)" in wb.sheetnames:
            ws = wb["BUDGET (INCOME)"]
            for r in range(5, 10):
                desc, amount = ws.cell(r, 2).value, ws.cell(r, 3).value
                if desc:
                    budgets[("INCOME", str(desc).strip())] = float(amount or 0)
        if "BUDGET (EXPENSES)" in wb.sheetnames:
            ws = wb["BUDGET (EXPENSES)"]
            for r in range(4, 50):
                desc, amount = ws.cell(r, 3).value, ws.cell(r, 4).value
                if desc and isinstance(amount, (int, float)):
                    budgets[("EXPENDITURE", str(desc).strip())] = float(amount or 0)
        wb.close()
    return list(dict.fromkeys(income)), list(dict.fromkeys(expenses)), budgets

INCOME_CATEGORY_MAP = {
    "Prophet Offering (Cash & Cheques)": "PROPHET OFFERING",
    "Bank transfers": "BANK",
    "Rent": "RENT",
    "Other dividends": "OTHERS (DIVIDENDS, ETC.)",
    "Forex": "FOREX",
}
INCOME_BUDGET_MAP = {
    "Prophet Offering (Cash & Cheques)": "PROPHET OFFERING",
    "Bank transfers": "BANK DIRECT TRANSFER",
    "Rent": "RENT",
    "Other dividends": "OTHERS (DIVIDENDS, ETC.)",
    "Forex": "FOREX INCOME",
}

def budget_category_key(txn_type, category):
    if txn_type == "INCOME":
        return INCOME_BUDGET_MAP.get(category, category)
    return category


def safe_sheet_name(name, used_names):
    """Return a unique, Excel-safe worksheet name (max 31 characters)."""
    cleaned = str(name).strip() or "Uncategorised"
    for char in ["\\", "/", "*", "?", ":", "[", "]"]:
        cleaned = cleaned.replace(char, "-")
    cleaned = cleaned[:31]
    candidate = cleaned
    counter = 2
    while candidate.casefold() in used_names:
        suffix = f" ({counter})"
        candidate = cleaned[:31-len(suffix)] + suffix
        counter += 1
    used_names.add(candidate.casefold())
    return candidate


def export_monthly_breakdowns(year, month, income_categories, expense_categories):
    """Export one worksheet per income/expense heading for the selected month."""
    from io import BytesIO
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter

    month_label = calendar.month_name[int(month)].upper()
    tx = load_transactions(year=year, month=month)
    workbook = Workbook()
    workbook.remove(workbook.active)
    used_names = set()
    all_categories = [("INCOME", c) for c in income_categories] + [("EXPENDITURE", c) for c in expense_categories]

    # Create every category tab, including categories with no transactions that month.
    for kind, category in all_categories:
        sheet_name = safe_sheet_name(category, used_names)
        ws = workbook.create_sheet(sheet_name)
        ws.append(["My Foundation"])
        ws.append([f"{category} : {month_label} {year}"])
        ws.append(["S/N", "DATE", "NAME / DESCRIPTION", "AMOUNT (N)", "PURPOSE", "NARRATION / DETAILS"])
        if not tx.empty:
            rows = tx[(tx["txn_type"] == kind) & (tx["category"] == category)].copy()
        else:
            rows = pd.DataFrame()
        if not rows.empty:
            for index, record in enumerate(rows.to_dict("records"), start=1):
                txn_day = pd.to_datetime(record["txn_date"]).date()
                name = str(record.get("detail_name") or "").strip()
                narration = str(record.get("narration") or "").strip()
                display_name = name or narration or category
                ws.append([index, txn_day, display_name, float(record["amount"]), record.get("purpose", ""), narration])
        else:
            ws.append([None, None, "No transactions recorded for this month", None, None, None])

        # Simple formatting inspired by the supplied Welfare breakdown workbook.
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
        ws["A1"].font = Font(bold=True, size=14)
        ws["A2"].font = Font(bold=True, size=12)
        for cell in ws[3]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="365F91")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        thin_grey = Side(style="thin", color="D9E1F2")
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=6):
            for cell in row:
                cell.border = Border(bottom=thin_grey)
        for row in ws.iter_rows(min_row=4, min_col=4, max_col=4):
            for cell in row:
                cell.number_format = '₦#,##0.00;[Red]-₦#,##0.00'
        for row in ws.iter_rows(min_row=4, min_col=2, max_col=2):
            for cell in row:
                cell.number_format = 'dd-mmm-yyyy'
        for col, width in {"A":8, "B":16, "C":34, "D":18, "E":28, "F":38}.items():
            ws.column_dimensions[col].width = width
        ws.freeze_panes = "A4"
        ws.auto_filter.ref = f"A3:F{ws.max_row}"

    # For fuel records, also create separate tabs named after each vehicle entered.
    if not tx.empty:
        fuel_rows = tx[(tx["txn_type"] == "EXPENDITURE") &
                       (tx["category"].str.contains("FUEL", case=False, na=False)) &
                       (tx["detail_name"].fillna("").str.strip() != "")]
        for vehicle in sorted(fuel_rows["detail_name"].astype(str).str.strip().unique(), key=str.casefold):
            vehicle_rows = fuel_rows[fuel_rows["detail_name"].astype(str).str.strip() == vehicle]
            ws = workbook.create_sheet(safe_sheet_name(vehicle, used_names))
            ws.append(["My Foundation"])
            ws.append([f"{vehicle} FUEL : {month_label} {year}"])
            ws.append(["S/N", "DATE", "NAME / DESCRIPTION", "AMOUNT (N)", "PURPOSE", "NARRATION / DETAILS"])
            for index, record in enumerate(vehicle_rows.to_dict("records"), start=1):
                ws.append([index, pd.to_datetime(record["txn_date"]).date(), vehicle,
                           float(record["amount"]), record.get("purpose", ""), record.get("narration", "")])
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)
            ws["A1"].font = Font(bold=True, size=14)
            ws["A2"].font = Font(bold=True, size=12)
            for cell in ws[3]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="365F91")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for row in ws.iter_rows(min_row=4, min_col=4, max_col=4):
                for cell in row:
                    cell.number_format = '₦#,##0.00;[Red]-₦#,##0.00'
            for row in ws.iter_rows(min_row=4, min_col=2, max_col=2):
                for cell in row:
                    cell.number_format = 'dd-mmm-yyyy'
            for col, width in {"A":8, "B":16, "C":34, "D":18, "E":28, "F":38}.items():
                ws.column_dimensions[col].width = width
            ws.freeze_panes = "A4"
            ws.auto_filter.ref = f"A3:F{ws.max_row}"

    output = BytesIO()
    workbook.save(output)
    return output.getvalue()


def load_transactions(year=None, month=None, txn_type=None, category=None):
    query = "SELECT * FROM transactions WHERE 1=1"
    params = []
    if year is not None:
        query += " AND substr(txn_date,1,4)=?"
        params.append(str(year))
    if month is not None:
        query += " AND substr(txn_date,6,2)=?"
        params.append(f"{int(month):02d}")
    if txn_type:
        query += " AND txn_type=?"
        params.append(txn_type)
    if category:
        query += " AND category=?"
        params.append(category)
    query += " ORDER BY txn_date, id"
    with connect() as con:
        rows = con.execute(query, params).fetchall()
    return pd.DataFrame([dict(r) for r in rows])

def add_transaction(txn_date, txn_type, category, amount, narration, detail_name, purpose):
    with connect() as con:
        con.execute(
            """INSERT INTO transactions
            (txn_date, txn_type, category, amount, narration, detail_name, purpose)
            VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (txn_date.isoformat(), txn_type, category, float(amount), narration.strip(),
             detail_name.strip(), purpose.strip())
        )

def export_transaction_file(df):
    from io import BytesIO
    output = BytesIO()
    cols = ["txn_date", "txn_type", "category", "amount", "narration", "detail_name", "purpose"]
    labels = ["Date", "Type", "Description / Category", "Amount (N)", "Narration", "Person / Vehicle", "Purpose"]
    df = df[cols].copy() if not df.empty else pd.DataFrame(columns=cols)
    df.columns = labels
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Transactions")
    return output.getvalue()

def export_monthly_summary(year, month, income_categories, expense_categories, budget_map):
    from io import BytesIO
    from openpyxl import Workbook
    month_label = calendar.month_name[int(month)].upper()
    tx = load_transactions(year=year, month=month)
    out = Workbook()
    ws = out.active
    ws.title = "Monthly Summary"
    ws.append([f"ACCOUNT SUMMARY — {month_label} {year}"])
    ws.append([])
    ws.append(["TYPE", "DESCRIPTION", "BUDGETED (ANNUAL)", "ACTUAL THIS MONTH", "YEAR-TO-DATE ACTUAL"])
    for kind, categories in [("INCOME", income_categories), ("EXPENDITURE", expense_categories)]:
        for category in categories:
            relevant = tx[(tx["txn_type"] == kind) & (tx["category"] == category)] if not tx.empty else pd.DataFrame()
            month_total = float(relevant["amount"].sum()) if not relevant.empty else 0
            ytd = load_transactions(year=year, txn_type=kind, category=category)
            ytd_total = float(ytd["amount"].sum()) if not ytd.empty else 0
            budget = budget_map.get((kind, budget_category_key(kind, category)), "")
            ws.append([kind, category, budget, month_total, ytd_total])
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:E{ws.max_row}"
    for col, width in {"A":16, "B":42, "C":22, "D":22, "E":22}.items():
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=4, min_col=3, max_col=5):
        for cell in row:
            cell.number_format = '₦#,##0.00;[Red]-₦#,##0.00'
    output = BytesIO()
    out.save(output)
    return output.getvalue()

init_db()
income_categories, expense_categories, budget_map = read_categories()

st.title("📒 Accounting & Account Summaries")
st.caption("Starter version • Enter transactions once, then review monthly and year-to-date totals.")

with st.sidebar:
    st.header("Navigation")
    page = st.radio("Choose a page", ["Enter transaction", "Transactions", "Monthly summary", "Export"])
    st.divider()
    st.caption("Your records are saved locally in accounting_records.db in this project folder.")

if page == "Enter transaction":
    st.subheader("Enter a transaction")
    col1, col2 = st.columns(2)
    with col1:
        txn_date = st.date_input("Transaction date", value=date.today())
        txn_type = st.selectbox("Transaction type", ["EXPENDITURE", "INCOME"])
        choices = expense_categories if txn_type == "EXPENDITURE" else income_categories
        category = st.selectbox("Description / category", choices)
        amount = st.number_input("Amount (₦)", min_value=0.0, step=1000.0, format="%.2f")
    with col2:
        narration = st.text_input("Narration / details", placeholder="What was this transaction for?")
        detail_name = st.text_input("Person or vehicle (optional)", placeholder="e.g. recipient name or vehicle name")
        purpose = st.text_input("Purpose (optional)", placeholder="e.g. school fees, medical, fuel")
        st.caption("For Welfare, enter the recipient in Person or vehicle. For fuel, enter the vehicle name. Each transaction is stored with its date and category.")
    if st.button("Save transaction", type="primary", use_container_width=True):
        if amount <= 0:
            st.error("Please enter an amount greater than zero.")
        else:
            add_transaction(txn_date, txn_type, category, amount, narration, detail_name, purpose)
            st.success("Transaction saved.")
            st.rerun()

elif page == "Transactions":
    st.subheader("Saved transactions")
    f1, f2, f3 = st.columns(3)
    with f1:
        year = st.number_input("Year", min_value=2000, max_value=2100, value=date.today().year, step=1)
    with f2:
        month = st.selectbox("Month", ["All months"] + [f"{i:02d} - {MONTHS[i-1]}" for i in range(1,13)])
    with f3:
        kind = st.selectbox("Type", ["All", "INCOME", "EXPENDITURE"])
    selected_month = None if month == "All months" else int(month[:2])
    df = load_transactions(year=int(year), month=selected_month, txn_type=None if kind == "All" else kind)
    if df.empty:
        st.info("No transactions match these filters yet.")
    else:
        st.dataframe(df[["id","txn_date","txn_type","category","amount","narration","detail_name","purpose"]],
                     use_container_width=True, hide_index=True)
        st.caption(f"{len(df)} transaction(s) shown. To correct a mistake, use the source record details and contact the app maintainer for the first version; editing/deletion will be added next.")

elif page == "Monthly summary":
    st.subheader("Monthly totals")
    c1, c2 = st.columns(2)
    with c1:
        year = st.number_input("Summary year", min_value=2000, max_value=2100, value=date.today().year, step=1)
    with c2:
        month = st.selectbox("Summary month", list(range(1,13)), format_func=lambda x: f"{x:02d} - {MONTHS[x-1]}")
    tx = load_transactions(year=int(year), month=int(month))
    income_total = float(tx.loc[tx["txn_type"] == "INCOME", "amount"].sum()) if not tx.empty else 0
    expense_total = float(tx.loc[tx["txn_type"] == "EXPENDITURE", "amount"].sum()) if not tx.empty else 0
    m1, m2, m3 = st.columns(3)
    m1.metric("Total income", f"₦{income_total:,.2f}")
    m2.metric("Total expenditure", f"₦{expense_total:,.2f}")
    m3.metric("Net movement", f"₦{income_total-expense_total:,.2f}")
    if not tx.empty:
        grouped = tx.groupby(["txn_type","category"], as_index=False)["amount"].sum()
        st.dataframe(grouped, use_container_width=True, hide_index=True)
        welfare = tx[(tx["txn_type"] == "EXPENDITURE") & (tx["category"].str.upper() == "WELFARE")]
        if not welfare.empty:
            st.info(f"Welfare total for {MONTHS[int(month)-1]}: ₦{welfare['amount'].sum():,.2f} — this is the sum of all Welfare transactions in that month.")
        fuel_names = tx[(tx["txn_type"] == "EXPENDITURE") & (tx["category"].str.upper().isin(["BENZ","CARINA","BMW","HIGHLANDER JEEP","NEW BLACK HIGHLANDER","WHITE PRADO JEEP","SIENNA"]))]
        if not fuel_names.empty:
            st.caption("Vehicle-related categories are shown separately using the descriptions in the uploaded expenditure template. We can refine which descriptions represent fuel after confirming the workbook's exact layout.")
    else:
        st.info("No transactions entered for this month yet.")

elif page == "Export":
    st.subheader("Export reports")
    c1, c2 = st.columns(2)
    with c1:
        year = st.number_input("Export year", min_value=2000, max_value=2100, value=date.today().year, step=1)
    with c2:
        month = st.selectbox("Export month", list(range(1,13)), format_func=lambda x: f"{x:02d} - {MONTHS[x-1]}")
    tx = load_transactions(year=int(year), month=int(month))
    st.write(f"Transactions available for export: **{len(tx)}**")
    st.download_button("Download monthly breakdown workbook (one sheet per heading)",
                       data=export_monthly_breakdowns(int(year), int(month), income_categories, expense_categories),
                       file_name=f"monthly_breakdowns_{year}_{int(month):02d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       help="Creates a worksheet for every income and expenditure heading, including headings with no transactions this month. Fuel transactions with a vehicle name also get a separate vehicle worksheet.")
    st.download_button("Download transactions (Excel)", data=export_transaction_file(tx),
                       file_name=f"transactions_{year}_{int(month):02d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("Download monthly summary (Excel)", data=export_monthly_summary(int(year), int(month), income_categories, expense_categories, budget_map),
                       file_name=f"monthly_summary_{year}_{int(month):02d}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.warning("This first version exports a clean summary workbook. The next build step is to populate your original four-sheet template directly and generate category/vehicle breakdown sheets in the same style as your Welfare sample.")

st.divider()
st.caption("Important: This is an early prototype. Test it with sample data first and keep backups of the original accounting workbooks.")
