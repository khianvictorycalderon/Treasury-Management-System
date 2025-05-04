import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook
import time
import threading

student_list = []
record_list = []
required_payments = {}

def load_student_status_data(excel_path):
    global student_list, record_list, required_payments
    student_list.clear()
    record_list.clear()
    required_payments.clear()

    try:
        wb = load_workbook(excel_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open Excel file: {e}")
        return

    if 'Student_Management' not in wb.sheetnames or 'Payment_Records' not in wb.sheetnames or 'Payment_Categories' not in wb.sheetnames:
        messagebox.showerror("Error", "Required sheets not found in Excel (Student_Management, Payment_Records, Payment_Categories).")
        return

    # Load students
    student_sheet = wb['Student_Management']
    for row in student_sheet.iter_rows(min_row=2, values_only=True):
        if row[0] is None: continue
        student_id, first_name, middle_initial, last_name = row
        full_name = f"{first_name} {last_name}"
        student_list.append({"id": student_id, "name": full_name})

    # Load required payments
    category_sheet = wb['Payment_Categories']
    for row in category_sheet.iter_rows(min_row=2, values_only=True):
        if row[0] is None: continue
        category, required_amount = row
        required_payments[category.lower()] = float(required_amount or 0)

    # Load payment records
    payment_sheet = wb['Payment_Records']
    for row in payment_sheet.iter_rows(min_row=2, values_only=True):
        # Ensure the row contains at least 4 values and safely handle any extra columns
        if len(row) >= 4:
            student_id, student_name, amount_paid, category = row[:4]  # Only take the first 4 columns
            if not student_id or not category:
                continue

            record_list.append({
                "student_id": student_id,
                "student_name": student_name,
                "amount_paid": float(amount_paid or 0),
                "payment_category": category.lower()
            })

def create_student_status_page(parent):
    frame = tk.Frame(parent)
    frame.pack(fill="both", expand=True)

    tk.Label(frame, text="Student Status Page", font=("Arial", 24)).pack(pady=10)

    # Table
    table_frame = tk.Frame(frame)
    table_frame.pack(expand=True, fill="both", padx=20, pady=10)

    columns = ["Student ID", "Student Name"] + [cat.title() for cat in required_payments.keys()]
    tree = ttk.Treeview(table_frame, columns=columns, show="headings")

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", stretch=True)

    vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(fill="both", expand=True)

    # Insert row data for student
    def insert_student_row(student):
        # Get all records for this student
        student_records = [r for r in record_list if r["student_id"] == student["id"]]

        payments_by_category = {cat: 0.0 for cat in required_payments.keys()}
        for record in student_records:
            cat = record["payment_category"]
            if cat in payments_by_category:
                payments_by_category[cat] += record["amount_paid"]

        row_data = [student["id"], student["name"]]
        for cat in required_payments.keys():
            paid = payments_by_category[cat]
            required = required_payments[cat]
            cell_value = f"{paid:.0f}/{required:.0f}" if required > 0 else ""
            row_data.append(cell_value)

        tree.insert("", "end", values=row_data)

    # Function to refresh the data every few seconds
    def refresh_data():
        load_student_status_data("userdata.xlsx")  # Load data from Excel
        tree.delete(*tree.get_children())  # Clear existing entries
        for student in student_list:
            insert_student_row(student)

        frame.after(10, refresh_data)  # Refresh every 5 seconds

    refresh_data()  # Start the data refresh

    return frame
