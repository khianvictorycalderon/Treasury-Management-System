import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook

#Global lists
student_list = []
record_list = []

#Load data from Excel
def load_student_status_data(excel_path):
    global student_list, record_list
    student_list.clear()
    record_list.clear()

    wb = load_workbook(excel_path)
    if 'Student_Management' not in wb.sheetnames or 'Payment_Records' not in wb.sheetnames:
        messagebox.showerror("Error", "Required sheets not found in Excel.")
        return

    student_sheet = wb['Student_Management']
    payment_sheet = wb['Payment_Records']

    #Load students
    for row in student_sheet.iter_rows(min_row=2, max_col=3, values_only=True):
        student_id, first_name, last_name = row
        full_name = f"{first_name} {last_name}"
        student_list.append({"id": student_id, "name": full_name})

    #Load relevant payments
    for row in payment_sheet.iter_rows(min_row=2, values_only=True):
        student_id, student_name, amount_paid, category = row
        if category.lower() in ["class fund", "project"]:
            record_list.append({
                "student_id": student_id,
                "student_name": student_name,
                "amount_paid": f"₱{amount_paid}",
                "payment_category": category
            })

#Build the GUI Page
def create_student_status_page(parent):
    frame = tk.Frame(parent)
    frame.pack(fill="both", expand=True)

    tk.Label(frame, text="Student Status Page", font=("Arial", 24)).pack(pady=10)

    #Search Bar
    def search_student():
        query = search_entry.get().strip().lower()
        if not query:
            messagebox.showerror("Input Error", "Please enter a student name to search.")
            return

        tree.delete(*tree.get_children())

        match_found = False
        for student in student_list:
            if query in student["name"].lower():
                insert_student_row(student)
                match_found = True

        if not match_found:
            messagebox.showinfo("No Match", "No student matched your search.")

    search_frame = tk.Frame(frame)
    search_frame.pack(pady=5, fill="x", padx=20)

    tk.Label(search_frame, text="Search Student:", font=("Arial", 12)).pack(side="left", padx=5)
    search_entry = tk.Entry(search_frame, width=50)
    search_entry.pack(side="left", fill="x", expand=True, padx=5)
    tk.Button(search_frame, text="Search", command=search_student).pack(side="left", padx=5)

    #Table
    table_frame = tk.Frame(frame)
    table_frame.pack(expand=True, fill="both", padx=20, pady=10)

    columns = ["Student Name", "Class Fund", "Project"]
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

    #Insert student row with Class Fund and Project totals
    def insert_student_row(student):
        student_records = [
            record for record in record_list if record["student_id"] == student["id"]
        ]

        class_fund_total = 0.0
        project_total = 0.0

        for record in student_records:
            amount_str = record["amount_paid"].replace("₱", "").replace(",", "")
            try:
                amount = float(amount_str)
            except ValueError:
                continue

            category = record["payment_category"].lower()
            if category == "class fund":
                class_fund_total += amount
            elif category == "project":
                project_total += amount

        tree.insert("", "end", values=[
            student["name"],
            f"₱{class_fund_total:.2f}" if class_fund_total else "",
            f"₱{project_total:.2f}" if project_total else ""
        ])

    #Insert all students initially
    for student in student_list:
        insert_student_row(student)

    return frame

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Student Status Page")
    root.geometry("800x500")

    load_student_status_data('userdata.xlsx')
    create_student_status_page(root)

    root.mainloop()
