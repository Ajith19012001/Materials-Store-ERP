import streamlit as st
import sqlite3
import pandas as pd
from datetime import date
from io import BytesIO
from fpdf import FPDF
import re

# ---------------- Database Initialization ----------------
conn = sqlite3.connect("civil_materials.db", check_same_thread=False)
c = conn.cursor()

c.execute("""CREATE TABLE IF NOT EXISTS purchase_orders (
    po_number TEXT PRIMARY KEY,
    supplier TEXT,
    date TEXT)""")

c.execute("""CREATE TABLE IF NOT EXISTS po_materials (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    po_number TEXT,
    material TEXT,
    unit TEXT,
    quantity REAL)""")

c.execute("""CREATE TABLE IF NOT EXISTS materials_received (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_number TEXT,
    po_number TEXT,
    material TEXT,
    quantity REAL,
    date TEXT)""")

c.execute("""CREATE TABLE IF NOT EXISTS materials_supplied (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    supply_number TEXT,
    po_number TEXT,
    contractor TEXT,
    material TEXT,
    quantity REAL,
    date TEXT)""")

conn.commit()

# ---------------- UI Setup ----------------
st.set_page_config(page_title="Civil Materials Management", layout="wide")
st.title("🏗️ Civil Materials Management")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["Purchase Orders", "Materials Received", "Materials Supplied", "Inventory", "Report", "Correction"]
)

# ---------------- Helper Functions ----------------
def get_next_receipt_number(po_number):
    c.execute(
        "SELECT receipt_number FROM materials_received WHERE po_number=? ORDER BY id DESC LIMIT 1",
        (po_number,)
    )
    last = c.fetchone()
    if last and last[0]:
        match = re.search(r'RE(\d+)$', last[0])
        num = int(match.group(1)) + 1 if match else 1
    else:
        num = 1
    return f"{po_number}-RE{num:02d}"

def get_next_supply_number(po_number):
    c.execute(
        "SELECT supply_number FROM materials_supplied WHERE po_number=? ORDER BY id DESC LIMIT 1",
        (po_number,)
    )
    last = c.fetchone()
    if last and last[0]:
        match = re.search(r'SU(\d+)$', last[0])
        num = int(match.group(1)) + 1 if match else 1
    else:
        num = 1
    return f"{po_number}-SU{num:02d}"

def export_excel(df, filename="report.xlsx"):
    buffer = BytesIO()
    # Uses xlsxwriter engine (Requires: pip install xlsxwriter)
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    st.download_button("⬇️ Download Excel", buffer.getvalue(), file_name=filename, mime="application/vnd.ms-excel")

def export_pdf(df, filename="report.pdf"):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=10)
    col_width = pdf.w / (len(df.columns) + 1)
    row_height = pdf.font_size
    for col in df.columns:
        pdf.cell(col_width, row_height*2, str(col), border=1)
    pdf.ln(row_height*2)
    for _, row in df.iterrows():
        for item in row:
            pdf.cell(col_width, row_height*2, str(item), border=1)
        pdf.ln(row_height*2)
    buffer = BytesIO(pdf.output(dest="S").encode("latin-1"))
    st.download_button("⬇️ Download PDF", buffer.getvalue(), file_name=filename, mime="application/pdf")

# ---------------- TAB 1: Purchase Orders ----------------
with tab1:
    st.subheader("Add Purchase Order")
    po_number = st.text_input("PO Number", key="po_number")
    supplier = st.text_input("Supplier", key="supplier")
    po_date = st.date_input("PO Date", date.today(), key="po_date")
    if st.button("Save PO", key="save_po"):
        if po_number and supplier:
            try:
                c.execute("INSERT INTO purchase_orders VALUES (?,?,?)", (po_number, supplier, str(po_date)))
                conn.commit()
                st.success("Purchase Order saved!")
            except sqlite3.IntegrityError:
                st.error("PO Number already exists!")
        else:
            st.warning("Please fill all fields.")

    st.subheader("Add Materials to PO")
    po_list = [row[0] for row in c.execute("SELECT po_number FROM purchase_orders").fetchall()]
    po_for_material = st.selectbox("PO Number (existing)", po_list, key="po_mat")
    mat_name = st.text_input("Material Name", key="mat_name")
    mat_unit = st.text_input("Unit (e.g. bags, tons)", key="mat_unit")
    mat_qty = st.number_input("Quantity as per PO", min_value=0.0, key="mat_qty")
    if st.button("Save Material", key="save_material"):
        if po_for_material and mat_name and mat_unit and mat_qty > 0:
            c.execute("INSERT INTO po_materials (po_number, material, unit, quantity) VALUES (?,?,?,?)",
                      (po_for_material, mat_name, mat_unit, mat_qty))
            conn.commit()
            st.success("Material added to PO!")
        else:
            st.warning("Please complete all material inputs with quantity greater than zero.")

# ---------------- TAB 2: Materials Received ----------------
with tab2:
    st.subheader("Record Materials Received")
    po_list = [row[0] for row in c.execute("SELECT po_number FROM purchase_orders").fetchall()]
    po_select = st.selectbox("Select PO Number", po_list, key="received_po")

    if po_select:
        receipt_number = get_next_receipt_number(po_select)
        st.text_input("Receipt Number (auto)", value=receipt_number, disabled=True)

        po_materials = c.execute("SELECT material, unit, quantity FROM po_materials WHERE po_number=?", (po_select,)).fetchall()
        rec_summary = c.execute("SELECT material, SUM(quantity) FROM materials_received WHERE po_number=? GROUP BY material", (po_select,)).fetchall()

        rows = []
        for i, (mat, unit, po_qty) in enumerate(po_materials, start=1):
            already_rec = next((r[1] for r in rec_summary if r[0] == mat), 0)
            balance = po_qty - already_rec
            rows.append({
                "Sl.No": i,
                "Material": mat,
                "Unit": unit,
                "PO Qty": po_qty,
                "Already Received": already_rec,
                "New Received Qty": 0.0,
                "Received Date": str(date.today()),
                "Balance Qty": balance
            })

        df = pd.DataFrame(rows)
        edited_df = st.data_editor(df, num_rows="fixed", key="received_editor",
                                   disabled=["Sl.No","Material","Unit","PO Qty","Already Received","Balance Qty"])

        if st.button("Save All Received Entries", key="save_received"):
            saved_any = False
            for _, row in edited_df.iterrows():
                new_qty = row["New Received Qty"]
                rec_date = row["Received Date"]
                if new_qty > 0:
                    c.execute("INSERT INTO materials_received (receipt_number, po_number, material, quantity, date) VALUES (?,?,?,?,?)",
                              (receipt_number, po_select, row["Material"], new_qty, str(rec_date)))
                    saved_any = True
            if saved_any:
                conn.commit()
                st.success(f"Entries saved under Receipt {receipt_number}!")
                st.rerun()

# ---------------- TAB 3: Materials Supplied ----------------
with tab3:
    st.subheader("Record Materials Supplied")
    po_list = [row[0] for row in c.execute("SELECT po_number FROM purchase_orders").fetchall()]
    po_select_sup = st.selectbox("Select PO Number", po_list, key="sup_po")

    if po_select_sup:
        supply_number = get_next_supply_number(po_select_sup)
        st.text_input("Supply Number (auto)", value=supply_number, disabled=True)

        po_materials = c.execute("SELECT material, unit, quantity FROM po_materials WHERE po_number=?", (po_select_sup,)).fetchall()
        rec_summary = c.execute("SELECT material, SUM(quantity) FROM materials_received WHERE po_number=? GROUP BY material", (po_select_sup,)).fetchall()
        sup_summary = c.execute("SELECT material, SUM(quantity) FROM materials_supplied WHERE po_number=? GROUP BY material", (po_select_sup,)).fetchall()

        rows = []
        for i, (mat, unit, po_qty) in enumerate(po_materials, start=1):
            rec_qty = next((r[1] for r in rec_summary if r[0] == mat), 0)
            sup_qty = next((s[1] for s in sup_summary if s[0] == mat), 0)
            stock = rec_qty - sup_qty
            rows.append({
                "Sl.No": i,
                "Material": mat,
                "Unit": unit,
                "PO Qty": po_qty,
                "Received Qty": rec_qty,
                "Already Supplied": sup_qty,
                "New Supply Qty": 0.0,
                "Supply Date": str(date.today()),
                "Contractor": "",
                "Stock Qty": stock
            })

        df = pd.DataFrame(rows)
        edited_df = st.data_editor(df, num_rows="fixed", key="supplied_editor",
                                   disabled=["Sl.No","Material","Unit","PO Qty","Received Qty","Already Supplied","Stock Qty"])

        if st.button("Save All Supplied Entries", key="save_supplied"):
            saved_any = False
            for _, row in edited_df.iterrows():
                new_qty = row["New Supply Qty"]
                sup_date = row["Supply Date"]
                contractor = row["Contractor"]
                if new_qty > 0:
                    if new_qty <= row["Stock Qty"]:
                        c.execute("INSERT INTO materials_supplied (supply_number, po_number, contractor, material, quantity, date) VALUES (?,?,?,?,?,?)",
                                  (supply_number, po_select_sup, contractor, row["Material"], new_qty, str(sup_date)))
                        saved_any = True
                    else:
                        st.error(f"Cannot supply {new_qty} of {row['Material']}. Only {row['Stock Qty']} available in stock!")
                        saved_any = False
                        break
            if saved_any:
                conn.commit()
                st.success(f"Entries saved under Supply ID {supply_number}!")
                st.rerun()

# ---------------- TAB 4: Inventory Status ----------------
with tab4:
    st.subheader("Current Physical Inventory Status")
    query = """
        SELECT 
            pm.po_number AS [PO Number],
            pm.material AS [Material Name],
            pm.unit AS [Unit],
            pm.quantity AS [PO Target Qty],
            COALESCE(mr.total_rec, 0) AS [Total Received],
            COALESCE(ms.total_sup, 0) AS [Total Supplied],
            (COALESCE(mr.total_rec, 0) - COALESCE(ms.total_sup, 0)) AS [Current On-Site Stock]
        FROM po_materials pm
        LEFT JOIN (SELECT po_number, material, SUM(quantity) as total_rec FROM materials_received GROUP BY po_number, material) mr 
            ON pm.po_number = mr.po_number AND pm.material = mr.material
        LEFT JOIN (SELECT po_number, material, SUM(quantity) as total_sup FROM materials_supplied GROUP BY po_number, material) ms 
            ON pm.po_number = ms.po_number AND pm.material = ms.material
    """
    inventory_df = pd.read_sql_query(query, conn)
    if not inventory_df.empty:
        st.dataframe(inventory_df, use_container_width=True, hide_index=True)
    else:
        st.info("No transaction data available to calculate inventory.")

# ---------------- TAB 5: Comprehensive Reports ----------------
with tab5:
    st.subheader("Data Export Center")
    report_type = st.selectbox("Select Report Category", ["Purchase Orders Master", "All Received Ledgers", "All Issued Ledgers"])
    
    if report_type == "Purchase Orders Master":
        rep_query = "SELECT po_number as [PO Number], material as [Material], unit as [Unit], quantity as [Quantity] FROM po_materials"
    elif report_type == "All Received Ledgers":
        rep_query = "SELECT receipt_number as [Receipt ID], po_number as [PO Number], material as [Material], quantity as [Qty Received], date as [Date] FROM materials_received"
    else:
        rep_query = "SELECT supply_number as [Supply ID], po_number as [PO Number], contractor as [Contractor], material as [Material], quantity as [Qty Issued], date as [Date] FROM materials_supplied"
    
    report_df = pd.read_sql_query(rep_query, conn)
    
    if not report_df.empty:
        st.dataframe(report_df, use_container_width=True, hide_index=True)
        col1, col2 = st.columns(2)
        with col1:
            export_excel(report_df, filename=f"{report_type.lower().replace(' ', '_')}.xlsx")
        with col2:
            export_pdf(report_df, filename=f"{report_type.lower().replace(' ', '_')}.pdf")
    else:
        st.info("No logs present for the selected profile.")

# ---------------- TAB 6: Corrections ----------------
with tab6:
    st.subheader("Transaction Entry Modifiers")
    mod_target = st.radio("Choose Records to Correct", ["Received Entries", "Supplied Entries"], horizontal=True)
    
    if mod_target == "Received Entries":
        logs_df = pd.read_sql_query("SELECT id, receipt_number, po_number, material, quantity, date FROM materials_received", conn)
        if not logs_df.empty:
            st.write("Modify quantities directly or clear rows:")
            edit_log = st.data_editor(logs_df, disabled=["id", "receipt_number", "po_number", "material"], key="corr_rec")
            if st.button("Apply Received Record Adjustments"):
                for _, row in edit_log.iterrows():
                    c.execute("UPDATE materials_received SET quantity = ?, date = ? WHERE id = ?", (row["quantity"], str(row["date"]), row["id"]))
                conn.commit()
                st.success("Received logs updated successfully.")
                st.rerun()
        else:
            st.info("No recorded receipts available.")
            
    else:
        logs_df = pd.read_sql_query("SELECT id, supply_number, po_number, contractor, material, quantity, date FROM materials_supplied", conn)
        if not logs_df.empty:
            st.write("Modify quantities directly or clear rows:")
            edit_log = st.data_editor(logs_df, disabled=["id", "supply_number", "po_number", "material"], key="corr_sup")
            if st.button("Apply Supplied Record Adjustments"):
                for _, row in edit_log.iterrows():
                    c.execute("UPDATE materials_supplied SET quantity = ?, contractor = ?, date = ? WHERE id = ?", (row["quantity"], row["contractor"], str(row["date"]), row["id"]))
                conn.commit()
                st.success("Supplied logs updated successfully.")
                st.rerun()
        else:
            st.info("No recorded supply entries available.")
