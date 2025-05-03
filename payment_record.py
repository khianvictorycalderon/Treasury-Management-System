import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from openpyxl import load_workbook

# Excel file path
excel_path = 'userdata.xlsx'
wb = load_workbook(excel_path)

# Access the sheets
student_sheet = wb['Student_Management']
payment_sheet = wb['Payment_Records']
# Initialize lists
record_list = []
category_list = ["Tuition", "Library Fee", "Miscellaneous"]
student_list = []

# Extract student data (IDs and Names) from Sheet1 (Student_Management)
for row in student_sheet.iter_rows(min_row=2, max_col=3, values_only=True):
    student_list.append({
        "id": row[0],  # Student ID
        "name": f"{row[1]} {row[2]}"  # Concatenating First Name and Last Name
    })

# Extract student IDs for Combobox
student_ids = [student["id"] for student in student_list]

# Function to create the Payment Record Page
def create_payment_record_page(parent):
    frame = tk.Frame(parent)
    frame.pack(fill="both", expand=True)

    tk.Label(frame, text="Payment Record Page", font=("Arial", 24)).pack(pady=10)

    # Input Fields
    form_frame = tk.Frame(frame)
    form_frame.pack(pady=5)

    # Combobox for Student IDs
    combobox_student_id = ttk.Combobox(form_frame, values=student_ids, state="readonly", width=28)
    combobox_student_id.grid(row=0, column=1, pady=2)
    tk.Label(form_frame, text="Student ID:").grid(row=0, column=0, sticky="e", padx=5, pady=2)

    # Entry for Amount Paid
    entry_amount = tk.Entry(form_frame, width=30)
    entry_amount.grid(row=1, column=1, pady=2)
    tk.Label(form_frame, text="Amount Paid:").grid(row=1, column=0, sticky="e", padx=5, pady=2)

    # Combobox for Payment Category
    combobox_category = ttk.Combobox(form_frame, values=category_list, state="readonly", width=28)
    combobox_category.grid(row=2, column=1, pady=2)
    tk.Label(form_frame, text="For:").grid(row=2, column=0, sticky="e", padx=5, pady=2)

    # Table for displaying records
    table_frame = tk.Frame(frame)
    table_frame.pack(expand=True, fill="both", padx=20, pady=10)

    columns = ("Student Name", "Amount Paid", "Payment Category", "Date & Time")
    tree = ttk.Treeview(table_frame, columns=columns, show="headings")
    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=150)

    vsb = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(expand=True, fill="both")

    def add_record():
        student_id = combobox_student_id.get().strip()
        amount_paid = entry_amount.get().strip()
        payment_category = combobox_category.get().strip()

        if not student_id or not amount_paid or not payment_category:
            messagebox.showerror("Input Error", "All fields must be filled out.")
            return

        try:
            float(amount_paid)
        except ValueError:
            messagebox.showerror("Input Error", "Amount must be a valid number.")
            return

        student_name = next((s["name"] for s in student_list if s["id"] == student_id), "Unknown")
        if student_name == "Unknown":
            messagebox.showerror("Input Error", "Student ID not found.")
            return

        # update UI
        date_time = datetime.now().strftime("%m/%d/%y, %I:%M %p")
        tree.insert("", "end", values=(student_name, f"₱{amount_paid}", payment_category, date_time))

        # immediately reload, append, and save Excel
        wb = load_workbook(excel_path)
        if "Payment_Records" in wb.sheetnames:
            sheet = wb["Payment_Records"]
        else:
            sheet = wb.create_sheet("Payment_Records")
            sheet.append(["Student ID","Student Name","Amount Paid","Payment Category","Date & Time"])
        sheet.append([student_id, student_name, amount_paid, payment_category, date_time])
        wb.save(excel_path)

        # clear inputs
        combobox_student_id.set('')
        combobox_category.set('')
        entry_amount.delete(0, tk.END)

    # Button to create a permanent record
    tk.Button(frame, text="Create Permanent Record", command=add_record).pack(pady=10)

    # Load existing records into the table (if any) from "Payment_Records"
    if payment_sheet:
        for row in payment_sheet.iter_rows(min_row=2, values_only=True):
            tree.insert("", "end", values=(row[1], f"₱{row[2]}", row[3], row[4]))

    return frame
    

# Launch the GUI
if __name__ == "__main__":
    root = tk.Tk()
    root.title("Payment Record Page")
    root.geometry("700x500")
    create_payment_record_page(root)
    root.mainloop()
