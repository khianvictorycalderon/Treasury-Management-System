from tkinter import *
from tkinter import messagebox  # For showing error messages
from student_management import create_student_management_page
from payment_category import create_payment_category_page
from payment_record import create_payment_record_page
from student_status import create_student_status_page
from bg_images import *
from bg_loader import load_bg_image
from credential import *

# Create root window
window = Tk()
window.title("Treasury Management System")
window.state("zoomed")
window.iconphoto(True, PhotoImage(file="Treasury_Management_Logo_Small_Icon.png"))

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
centerFrame = Frame(loginPage)
centerFrame.grid(row=0, column=0, sticky="n", pady=240)

# Make loginPage expand and center the content
loginPage.grid_rowconfigure(0, weight=1)
loginPage.grid_columnconfigure(0, weight=1)

# Add login widgets inside centerFrame
loginLabel = Label(centerFrame, text="Log In", font=("Arial", 40, "bold"))
loginLabel.grid(row=0, column=0, pady=(0, 30))

userName = Entry(centerFrame, font=("Arial", 15))
userName.grid(row=1, column=0, pady=10)

passWord = Entry(centerFrame, font=("Arial", 15), show="*")
passWord.grid(row=2, column=0, pady=10)

# Define login validation function
def validate_login():
    entered_username = userName.get()
    entered_password = passWord.get()
    
    if entered_username == CORRECT_USERNAME and entered_password == CORRECT_PASSWORD:
        changePage("student_management_page")  # Move to the next page on successful login
    else:
        messagebox.showerror("Login Failed", "Incorrect username or password. Please try again.")  # Show error message

loginButton = Button(centerFrame, text="Log In", width=20, command=validate_login)
loginButton.grid(row=3, column=0, pady=20)

changePage("login_page")

window.mainloop()
