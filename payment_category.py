from tkinter import *
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os

category_list = []

def export_to_excel():
    if not category_list:
        messagebox.showwarning("No Data", "There are no categories to export.")
        return

    # open or create workbook
    if os.path.exists('userdata.xlsx'):
        workbook = load_workbook('userdata.xlsx')
    else:
        workbook = Workbook()

    # get—or make—a specific sheet called "Payment_Categories"
    if "Payment_Categories" in workbook.sheetnames:
        sheet = workbook["Payment_Categories"]
        
    else:
        sheet = workbook.create_sheet("Payment_Categories")
        sheet = workbook.active
        sheet.append(["Payment Category Name", "Required Fund (Per Student)"])

    # append your current list
    for name, fund in category_list:
        sheet.append([name, fund])

    workbook.save("userdata.xlsx")
    messagebox.showinfo("Success", "Exported to 'userdata.xlsx'")

def create_payment_category_page(parent):
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

    btn_add = Button(frame, text="Add Category", font=("Arial", 14), command=lambda: add_category())
    btn_add.grid(row=2, column=0, columnspan=2, pady=(10, 20))

    btn_export = Button(frame, text="Export to Excel", font=("Arial", 14), command=export_to_excel)
    btn_export.grid(row=2, columnspan= 2, pady=(10, 20), sticky="e")

    frame_table = Frame(frame)
    frame_table.grid(row=3, column=0, columnspan=2, pady=10)
    

    def refresh_table():
        for widget in frame_table.winfo_children():
            widget.destroy()

        Label(frame_table, text="Payment Category Name", font=("Arial", 14, "bold"), borderwidth=2, relief="solid", width=30).grid(row=0, column=0, padx=2, pady=2)
        Label(frame_table, text="Required Fund", font=("Arial", 14, "bold"), borderwidth=2, relief="solid", width=20).grid(row=0, column=1, padx=2, pady=2)

        sorted_categories = sorted(category_list, key=lambda x: x[0])

        for i, (name, fund) in enumerate(sorted_categories, start=1):
            name_var = StringVar(value=name)
            fund_var = StringVar(value=str(fund))

            entry_name_edit = Entry(frame_table, textvariable=name_var, font=("Arial", 12), width=30, state="disabled")
            entry_fund_edit = Entry(frame_table, textvariable=fund_var, font=("Arial", 12), width=20, state="disabled")

            entry_name_edit.grid(row=i, column=0, padx=2, pady=2)
            entry_fund_edit.grid(row=i, column=1, padx=2, pady=2)

            def make_edit_handler(index=i-1, ename=entry_name_edit, efund=entry_fund_edit,
                                  nv=name_var, fv=fund_var, btn=None):
                def handler():
                    if ename["state"] == "disabled":
                        # Enable editing
                        ename.config(state="normal")
                        efund.config(state="normal")
                        btn.config(text="Save")
                    else:
                        # Save changes
                        new_name = nv.get().strip()
                        new_fund = fv.get().strip()
                        if not new_name or not new_fund:
                            messagebox.showwarning("Input Error", "Fields cannot be empty.")
                            return
                        try:
                            fund_amount = int(new_fund)
                        except ValueError:
                            messagebox.showerror("Input Error", "Required Fund must be a number.")
                            return

                        # Check for name conflict
                        for j, (cat_name, _) in enumerate(category_list):
                            if j != index and cat_name == new_name:
                                messagebox.showerror("Duplicate", "Another category with that name already exists.")
                                return

                        category_list[index] = (new_name, fund_amount)
                        refresh_table()
                return handler

            # Create button separately so we can reference it in the handler
            btn_edit = Button(frame_table, text="Edit", font=("Arial", 10))
            btn_edit.grid(row=i, column=2, padx=2, pady=2)
            btn_edit.config(command=make_edit_handler(btn=btn_edit))

            Button(frame_table, text="Delete", font=("Arial", 10),
                   command=lambda n=name: delete_category(n)).grid(row=i, column=3, padx=2, pady=2)

    def add_category():
        name = entry_name.get().strip()
        fund = entry_fund.get().strip()

        if not name or not fund:
            messagebox.showwarning("Input Error", "Please fill out all fields.")
            return

        try:
            fund_amount = int(fund)
        except ValueError:
            messagebox.showerror("Input Error", "Required Fund must be a number.")
            return

        if any(cat[0] == name for cat in category_list):
            messagebox.showwarning("Duplicate", "Category already exists.")
            return

        category_list.append((name, fund_amount))
        entry_name.delete(0, END)
        entry_fund.delete(0, END)
        refresh_table()

    def delete_category(name):
        global category_list
        category_list[:] = [cat for cat in category_list if cat[0] != name]

        if os.path.exists('userdata.xlsx'):
            try:
                workbook = load_workbook('userdata.xlsx')
                
                if "Payment_Categories" in workbook.sheetnames:
                    sheet = workbook["Payment_Categories"]

                    for row in range(2, sheet.max_row + 1):
                        cell_value = sheet[f"A{row}"].value
                        if cell_value == name:
                            sheet.delete_rows(row)
                            break

                    workbook.save('userdata.xlsx')
            except Exception as e:
                messagebox.showerror("Error", f"Could not update Excel file: {e}")

    refresh_table()
    return outer_frame
