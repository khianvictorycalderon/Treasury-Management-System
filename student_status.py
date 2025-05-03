import tkinter as tk
from tkinter import ttk, messagebox

student_list=[]
record_list=[]

def create_student_status_page(parent):
    
    global student_list
    global record_list
    
    frame = tk.Frame(parent)
    frame.pack(fill="both", expand=True)

    tk.Label(frame, text="Student Status Page", font=("Arial", 24)).pack(pady=10)

    # Search Function
    def search_student():
        query = search_entry.get().strip().lower()
        if not query:
            messagebox.showerror("Input Error", "Please enter a student name to search.")
            return

        tree.delete(*tree.get_children())

        match_found = False
        for student in student_list:
            if query in student["name"].lower():
                # Find corresponding records for this student
                student_records = [
                    record for record in record_list if record["student_id"] == student["id"]
                ]
                # Insert student and their records into the tree
                record_values = [student["name"]] + [
                    f"{record['amount_paid']} ({record['payment_category']})"
                    for record in student_records
                ]
                tree.insert("", "end", values=record_values)
                match_found = True

        if not match_found:
            messagebox.showinfo("No Match", "No student matched your search.")

    # Search Bar
    search_frame = tk.Frame(frame)
    search_frame.pack(pady=5, fill="x", padx=20)

    tk.Label(search_frame, text="Search Student:", font=("Arial", 12)).pack(side="left", padx=5)
    search_entry = tk.Entry(search_frame, width=50)
    search_entry.pack(side="left", fill="x", expand=True, padx=5)
    tk.Button(search_frame, text="Search", command=search_student).pack(side="left", padx=5)

    # Table for Student Status
    table_frame = tk.Frame(frame)
    table_frame.pack(expand=True, fill="both", padx=20, pady=10)

    columns = ["Student Name", "Class fund (Jan)", "Class fund (Feb)", "Project"]
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

    # Insert Initial Data from `student_list` and `record_list`
    for student in student_list:
        student_records = [
            record for record in record_list if record["student_id"] == student["id"]
        ]
        record_values = [student["name"]] + [
            f"{record['amount_paid']} ({record['payment_category']})"
            for record in student_records
        ]
        tree.insert("", "end", values=record_values)

    return frame
