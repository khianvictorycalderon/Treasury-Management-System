from tkinter import *
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os

category_list = []
EXCEL_FILE = 'userdata.xlsx'
CATEGORY_SHEET = 'Payment_Categories'

def load_from_excel():
    global category_list
    category_list.clear()
    if os.path.exists(EXCEL_FILE):
        try:
            wb = load_workbook(EXCEL_FILE)
            if CATEGORY_SHEET in wb.sheetnames:
                sheet = wb[CATEGORY_SHEET]

                for row in sheet.iter_rows(min_row=1, values_only=True):
                    if row[0] is None:
                        continue
                    name = str(row[0]).strip()
                    try:
                        fund = float(row[1])
                    except Exception:
                        fund = 0
                    if not any(cat[0].lower() == name.lower() for cat in category_list):
                        category_list.append((name, fund))
        except Exception as e:
            messagebox.showerror("Error Loading Excel", f"Could not load Excel: {e}")

def export_to_excel():
    try:
        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
        else:
            wb = Workbook()

        if CATEGORY_SHEET in wb.sheetnames:
            sheet = wb[CATEGORY_SHEET]
        else:
            sheet = wb.create_sheet(CATEGORY_SHEET)
            sheet.append(['Payment Category Name', 'Required Fund (Per Student)'])

        if sheet.max_row > 1:
            sheet.delete_rows(2, sheet.max_row - 1)

        for name, fund in category_list:
            sheet.append([name, fund])

        wb.save(EXCEL_FILE)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to export to Excel: {e}")

def create_payment_category_page(parent):
    global category_list
    outer_frame = Frame(parent)
    outer_frame.pack(fill="both", expand=True)
    outer_frame.grid_rowconfigure(0, weight=1)
    outer_frame.grid_columnconfigure(0, weight=1)

    frame = Frame(outer_frame)
    frame.grid(row=0, column=0)

    Label(frame, text="Payment Category Page", font=("Arial", 40, "bold")).grid(row=0, column=0, columnspan=2, pady=(30, 20))

    input_frame = Frame(frame)
    input_frame.grid(row=1, column=0, columnspan=2, pady=10)

    Label(input_frame, text="Payment Category Name:", font=("Arial", 14)).grid(row=0, column=0, padx=10, pady=5, sticky="e")
    entry_name = Entry(input_frame, width=30, font=("Arial", 14))
    entry_name.grid(row=0, column=1, padx=10, pady=5)

    Label(input_frame, text="Required Fund (Per Student):", font=("Arial", 14)).grid(row=1, column=0, padx=10, pady=5, sticky="e")
    entry_fund = Entry(input_frame, width=30, font=("Arial", 14))
    entry_fund.grid(row=1, column=1, padx=10, pady=5)

    frame_table = Frame(frame)
    frame_table.grid(row=3, column=0, columnspan=2, pady=10)

    def refresh_table():
        for widget in frame_table.winfo_children():
            widget.destroy()

        Label(frame_table, text="Payment Category Name", font=("Arial", 14, "bold"),
              borderwidth=2, relief="solid", width=30).grid(row=0, column=0, padx=2, pady=2)
        Label(frame_table, text="Required Fund", font=("Arial", 14, "bold"),
              borderwidth=2, relief="solid", width=20).grid(row=0, column=1, padx=2, pady=2)
        Label(frame_table, text="", width=10).grid(row=0, column=2)
        Label(frame_table, text="", width=10).grid(row=0, column=3)

        for i, (name, fund) in enumerate(category_list):
            Label(frame_table, text=name, font=("Arial", 12), width=30, anchor="w",
                  borderwidth=1, relief="solid").grid(row=i + 1, column=0, padx=2, pady=2)
            Label(frame_table, text=str(fund), font=("Arial", 12), width=20, anchor="w",
                  borderwidth=1, relief="solid").grid(row=i + 1, column=1, padx=2, pady=2)

            def delete_handler(idx=i):
                if messagebox.askyesno("Confirm Delete", f"Delete category '{category_list[idx][0]}'?"):
                    category_list.pop(idx)
                    export_to_excel()
                    refresh_table()

            Button(frame_table, text="Delete", font=("Arial", 10), command=delete_handler).grid(row=i + 1, column=2, padx=2, pady=2)

    def add_category():
        name = entry_name.get().strip()
        fund = entry_fund.get().strip()

        if not name or not fund:
            messagebox.showerror("Input Error", "Both fields are required.")
            return

        try:
            fund_value = float(fund)
            if fund_value <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Input Error", "Required fund must be a positive number.")
            return

        if any(cat[0].lower() == name.lower() for cat in category_list):
            messagebox.showwarning("Duplicate Entry", "This category already exists.")
            return

        category_list.append((name, fund_value))
        export_to_excel()
        refresh_table()

        entry_name.delete(0, END)
        entry_fund.delete(0, END)

    btn_add = Button(frame, text="Add Category", font=("Arial", 14), command=add_category)
    btn_add.grid(row=2, column=0, columnspan=2, pady=(10, 20))

    load_from_excel()
    refresh_table()

    def auto_refresh():
        load_from_excel()
        refresh_table()
        outer_frame.after(1000, auto_refresh)

    auto_refresh()

    return outer_frame
