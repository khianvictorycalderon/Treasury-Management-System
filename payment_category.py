from tkinter import *
from tkinter import messagebox
from openpyxl import Workbook, load_workbook
import os

category_list = []

def load_from_excel():
    global category_list
    category_list.clear()
    if os.path.exists('userdata.xlsx'):
        try:
            wb = load_workbook('userdata.xlsx')
            if 'Payment_Categories' in wb.sheetnames:
                sheet = wb['Payment_Categories']
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    if row[0] is None:
                        continue
                    name = str(row[0]).strip()
                    try:
                        fund = int(row[1])
                    except Exception:
                        fund = 0
                    # avoid duplicates when loading
                    if not any(cat[0].lower() == name.lower() for cat in category_list):
                        category_list.append((name, fund))
        except Exception as e:
            messagebox.showerror("Error Loading Excel", f"Could not load Excel: {e}")

def export_to_excel():
    if not category_list:
        # If no category, optionally clear excel sheet
        if os.path.exists('userdata.xlsx'):
            try:
                wb = load_workbook('userdata.xlsx')
                if 'Payment_Categories' in wb.sheetnames:
                    sheet = wb['Payment_Categories']
                    sheet.delete_rows(2, sheet.max_row - 1 if sheet.max_row > 1 else 0)
                    wb.save('userdata.xlsx')
                return
            except Exception as e:
                messagebox.showerror("Error", f"Failed to clear Excel: {e}")
                return
        return

    try:
        if os.path.exists('userdata.xlsx'):
            workbook = load_workbook('userdata.xlsx')
        else:
            workbook = Workbook()

        if 'Payment_Categories' in workbook.sheetnames:
            sheet = workbook['Payment_Categories']
            if sheet.max_row > 1:
                sheet.delete_rows(2, sheet.max_row -1)
        else:
            sheet = workbook.create_sheet('Payment_Categories')
            sheet.append(['Payment Category Name', 'Required Fund (Per Student)'])

        for name, fund in category_list:
            sheet.append([name, fund])

        workbook.save('userdata.xlsx')

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

    btn_add = Button(frame, text="Add Category", font=("Arial", 14))
    btn_add.grid(row=2, column=0, columnspan=2, pady=(10, 20))

    frame_table = Frame(frame)
    frame_table.grid(row=3, column=0, columnspan=2, pady=10)

    edit_mode_index = [None]

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
            if edit_mode_index[0] == i:
                name_var = StringVar(value=name)
                fund_var = StringVar(value=str(fund))

                entry_name_edit = Entry(frame_table, textvariable=name_var, font=("Arial", 12), width=30)
                entry_fund_edit = Entry(frame_table, textvariable=fund_var, font=("Arial", 12), width=20)

                entry_name_edit.grid(row=i +1, column=0, padx=2, pady=2)
                entry_fund_edit.grid(row=i +1, column=1, padx=2, pady=2)

                def save_handler(idx=i, nv=name_var, fv=fund_var):
                    new_name = nv.get().strip()
                    new_fund = fv.get().strip()
                    if not new_name or not new_fund:
                        messagebox.showwarning("Input Error", "Please fill all fields.")
                        return

                    try:
                        fund_amount = int(new_fund)
                    except ValueError:
                        messagebox.showerror("Input Error", "Required Fund must be a number.")
                        return

                    # Check duplicate (case insensitive) excluding current
                    for j, (cat_name, _) in enumerate(category_list):
                        if j != idx and cat_name.lower() == new_name.lower():
                            messagebox.showerror("Duplicate", "Another category with that name already exists.")
                            return

                    category_list[idx] = (new_name, fund_amount)
                    edit_mode_index[0] = None
                    export_to_excel()
                    refresh_table()

                btn_save = Button(frame_table, text="Save", font=("Arial", 10), command=save_handler)
                btn_save.grid(row=i +1, column=2, padx=2, pady=2)

                btn_cancel = Button(frame_table, text="Cancel", font=("Arial", 10),
                                    command=lambda: cancel_edit())
                btn_cancel.grid(row=i +1, column=3, padx=2, pady=2)

            else:
                label_name = Label(frame_table, text=name, font=("Arial", 12), width=30, anchor="w", borderwidth=1, relief="solid")
                label_fund = Label(frame_table, text=str(fund), font=("Arial", 12), width=20, anchor="w", borderwidth=1, relief="solid")

                label_name.grid(row=i +1, column=0, padx=2, pady=2)
                label_fund.grid(row=i +1, column=1, padx=2, pady=2)

                def edit_handler(idx=i):
                    if edit_mode_index[0] is None:
                        edit_mode_index[0] = idx
                        refresh_table()
                    else:
                        messagebox.showinfo("Editing in Progress", "Finish editing the current row before editing another.")

                def delete_handler(idx=i):
                    if messagebox.askyesno("Confirm Delete", f"Delete category '{category_list[idx][0]}'?"):
                        if edit_mode_index[0] == idx:
                            edit_mode_index[0] = None
                        category_list.pop(idx)
                        export_to_excel()
                        refresh_table()

                btn_edit = Button(frame_table, text="Edit", font=("Arial", 10), command=edit_handler)
                btn_edit.grid(row=i +1, column=2, padx=2, pady=2)

                btn_delete = Button(frame_table, text="Delete", font=("Arial", 10), command=delete_handler)
                btn_delete.grid(row=i +1, column=3, padx=2, pady=2)

    def cancel_edit():
        edit_mode_index[0] = None
        refresh_table()

    def add_category():
        if edit_mode_index[0] is not None:
            messagebox.showinfo("Editing in Progress", "Please finish current editing before adding a new category.")
            return

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

        # Prevent duplicates, case insensitive
        if any(cat[0].lower() == name.lower() for cat in category_list):
            messagebox.showwarning("Duplicate", "Category already exists.")
            return

        category_list.append((name, fund_amount))
        entry_name.delete(0, END)
        entry_fund.delete(0, END)
        export_to_excel()
        refresh_table()

    btn_add.config(command=add_category)

    load_from_excel()
    refresh_table()
    return outer_frame