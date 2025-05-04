import tkinter as tk
from tkinter import messagebox
from openpyxl import load_workbook

# Global lists
student_list = []
record_list = []

# Load data from Excel
def load_payment_data(excel_path):
    global student_list, record_list
    student_list.clear()
    record_list.clear()

    try:
        wb = load_workbook(excel_path)

        # Ensure required sheets exist; create if missing
        if "Student_Management" not in wb.sheetnames:
            student_sheet = wb.create_sheet("Student_Management")
            student_sheet.append(["ID", "First Name", "Middle Initial", "Last Name"])
        if "Payment_Records" not in wb.sheetnames:
            payment_sheet = wb.create_sheet("Payment_Records")
            payment_sheet.append(["Student ID", "Student Name", "Amount Paid", "Category"])

        student_sheet = wb["Student_Management"]
        payment_sheet = wb["Payment_Records"]

        # Load students
        for row in student_sheet.iter_rows(min_row=2, values_only=True):
            if row[0] and row[1] and row[3]:  # ID, First Name, Last Name
                student_id, first_name, _, last_name = row
                full_name = f"{first_name} {last_name}"
                student_list.append({"id": student_id, "name": full_name})

        # Load payment records
        for row in payment_sheet.iter_rows(min_row=2, values_only=True):
            student_id, student_name, amount_paid, category = row
            if student_id and student_name and amount_paid and category:
                record_list.append({
                    "student_id": student_id,
                    "student_name": student_name,
                    "amount_paid": f"₱{amount_paid}",
                    "payment_category": category
                })

        # Save back to file in case we added sheets
        wb.save(excel_path)

    except FileNotFoundError:
        messagebox.showerror("File Error", f"'{excel_path}' not found.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while loading data:\n{e}")

# GUI page creation
def create_payment_record_page(parent):
    frame = tk.Frame(parent)
    frame.pack(fill="both", expand=True)

    tk.Label(frame, text="Payment Records", font=("Arial", 24)).pack(pady=10)

    # Table
    table_frame = tk.Frame(frame)
    table_frame.pack(expand=True, fill="both", padx=20, pady=10)

    tree = tk.ttk.Treeview(table_frame, columns=["ID", "Name", "Amount", "Category"], show="headings")

    for col in ["ID", "Name", "Amount", "Category"]:
        tree.heading(col, text=col)
        tree.column(col, anchor="center", width=150)

    vsb = tk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)

    tree.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")

    # Insert records
    for record in record_list:
        tree.insert("", "end", values=[
            record["student_id"],
            record["student_name"],
            record["amount_paid"],
            record["payment_category"]
        ])

    return frame

# For testing standalone
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Payment Record Page")
    root.geometry("700x500")

    load_payment_data("userdata.xlsx")
    create_payment_record_page(root)

    root.mainloop()
