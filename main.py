from tkinter import *
from tkinter import messagebox  # For showing error messages
from student_management import create_student_management_page
from payment_category import create_payment_category_page
from payment_record import create_payment_record_page
from student_status import create_student_status_page
from bg_images import *
from bg_loader import load_bg_image
from credential import *
from icon import *

BLUE = "#4190F3"
YELLOW = "#F5B940"
DARK_BLUE = "#17355A"
GREEN = "#009820"
WHITE = "#FFFFFF"
BG_ENTRY = "#F7FAFC"

# Create root window
window = Tk()
window.title("Treasury Management System")
window.state("zoomed")
raw_base64 = ICON_BASE_64.split(",", 1)[1]
icon = PhotoImage(data=raw_base64)
window.iconphoto(True, icon)

# Create pages as frames
loginPage = Frame(window)
loginPage.pack()

studentManagementPage = create_student_management_page(window)
paymentCategoryPage = create_payment_category_page(window)
paymentRecordPage = create_payment_record_page(window)
studentStatusPage = create_student_status_page(window)

currentPage = "login_page"  # Default startup page
load_bg_image(loginPage, LOGIN_PAGE_BACKGROUND_IMAGE)

def changePage(pageName):
    global paymentCategoryPage  # Make sure paymentCategoryPage is global for reference
    
    # Hide all pages first
    loginPage.pack_forget()
    studentManagementPage.pack_forget()
    paymentCategoryPage.pack_forget()
    paymentRecordPage.pack_forget()
    studentStatusPage.pack_forget()
    
    global currentPage
    
    if pageName == "login_page":
        userName.delete(0, END)
        passWord.delete(0, END)
        currentPage = "login_page"
        loginPage.pack(fill=BOTH, expand=TRUE) 
    elif pageName == "student_management_page":
        currentPage = "student_management_page"
        studentManagementPage.pack(fill=BOTH, expand=TRUE)
    elif pageName == "payment_category_page":
        currentPage = "payment_category_page"
        
        if 'paymentCategoryPage' in globals():  # Check if the page already exists
            paymentCategoryPage.destroy()  # Destroy the existing page
        
        # Create the page again with the updated category list
        paymentCategoryPage = create_payment_category_page(window)
        paymentCategoryPage.pack(fill=BOTH, expand=TRUE)
    elif pageName == "payment_record_page":
        currentPage = "payment_record_page"
        paymentRecordPage.pack(fill=BOTH, expand=TRUE)
    elif pageName == "student_status_page":
        currentPage = "student_status_page"
        studentStatusPage.pack(fill=BOTH, expand=TRUE)
    if currentPage == "login_page":
        window.config(menu="")  # Remove menu from login page
    else:
        menu = Menu(window, font=("Arial", 20))
        window.config(menu=menu)
        menu.add_cascade(label="Student Management", command=lambda: changePage("student_management_page"))
        menu.add_cascade(label="Payment Category", command=lambda: changePage("payment_category_page"))
        menu.add_cascade(label="Payment Record", command=lambda: changePage("payment_record_page"))
        menu.add_cascade(label="Student Status", command=lambda: changePage("student_status_page"))
        menu.add_cascade(label="Log Out", command=lambda: changePage("login_page"))

# Create a center container inside loginPage
centerFrame = Frame(loginPage, bg=WHITE, bd=2, relief=RIDGE, padx=70, pady=40)
centerFrame.place(relx=0.5, rely=0.5, anchor=CENTER)

# Make loginPage expand and center the content
loginPage.grid_rowconfigure(0, weight=1)
loginPage.grid_columnconfigure(0, weight=1)

# Add login widgets inside centerFrame
loginLabel = Label(centerFrame, text="Log In", font=("Segoe UI", 36, "bold"), fg=BLUE, bg=WHITE)
loginLabel.grid(row=0, column=0, columnspan=2, pady=(0, 30))

Label(centerFrame, text="Username", font=("Segoe UI", 14, "bold"), fg=DARK_BLUE, bg=WHITE).grid(row=1, column=0, sticky="e", padx=10, pady=8)
userName = Entry(centerFrame, font=("Segoe UI", 14), fg=DARK_BLUE, bg=BG_ENTRY,
                 highlightthickness=2, highlightbackground=DARK_BLUE, relief=FLAT, bd=0, width=22)
userName.grid(row=1, column=1, pady=8, ipady=4)

Label(centerFrame, text="Password", font=("Segoe UI", 14, "bold"), fg=DARK_BLUE, bg=WHITE).grid(row=2, column=0, sticky="e", padx=10, pady=8)
passWord = Entry(centerFrame, font=("Segoe UI", 14), fg=DARK_BLUE, bg=BG_ENTRY, show="*",
                 highlightthickness=2, highlightbackground=DARK_BLUE, relief=FLAT, bd=0, width=22)
passWord.grid(row=2, column=1, pady=8, ipady=4)

def on_login_enter(e): loginButton.config(bg="#007A17")
def on_login_leave(e): loginButton.config(bg=BLUE)

# Define login validation function
def validate_login():
    entered_username = userName.get()
    entered_password = passWord.get()
    
    if entered_username == CORRECT_USERNAME and entered_password == CORRECT_PASSWORD:
        changePage("student_management_page")  # Move to the next page on successful login
    else:
        messagebox.showerror("Login Failed", "Incorrect username or password. Please try again.")  # Show error message

loginButton = Button(centerFrame, text="Log In", font=("Segoe UI", 14, "bold"),
                     activebackground=GREEN, activeforeground=WHITE,
                     bd=0, relief=FLAT, cursor="hand2", width=18, bg="#1A7FB7", fg="white", command=validate_login)
loginButton.grid(row=3, column=0, columnspan=2, pady=(25, 10))
loginButton.bind("<Enter>", on_login_enter)
loginButton.bind("<Leave>", on_login_leave)

# --- Forgot Password (hided) ---
def on_forgot_enter(e): forgotButton.config(bg=WHITE)
def on_forgot_leave(e): forgotButton.config(bg=WHITE)

forgotButton = Button(centerFrame, text="Forgot Password?", font=("Segoe UI", 12, "bold"),
                      bg=WHITE, fg=WHITE, activebackground=WHITE, activeforeground=WHITE,
                      bd=0, relief=FLAT, width=18,
                      command=lambda: messagebox.showinfo("Forgot Password", "Please contact admin."))
forgotButton.grid(row=4, column=0, columnspan=2, pady=(0, 5))
forgotButton.bind("<Enter>", on_forgot_enter)
forgotButton.bind("<Leave>", on_forgot_leave)

changePage("login_page")

window.mainloop()
