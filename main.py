from tkinter import *
from tkinter import messagebox
import random
import subprocess
import json
# import pyperclip

DEEP_BLUE = "#004080"
ACCENT_ORANGE = "#ffa500"
STEEL_GREY = "#9a9a9a"
CHARCOAL = "#333333"


# Back-----------------------------------------------------Back
def run_powershell_command(powershell_script, expect_json=True):
    """ This method meant to run powershell commands, it takes powershell command as a variable, if JSON output is needed(like data of a user) expect_json
    equals True, if no json needed, set it to False"""
    try:
        result = subprocess.run(["powershell", "-Command", powershell_script], capture_output=True, text=True,
                                check=True)

        if expect_json:
            try:
                return json.loads(result.stdout)
            except json.JSONDecodeError:
                # Return None if JSON parsing fails
                messagebox.showinfo(title="JSONDecode ERROR", message=result.stdout)
                print("Error decoding JSON. Output:", result.stdout)
                return None
        else:
            return result.stdout
    except subprocess.CalledProcessError as e:
        messagebox.showerror(title="PowerShell Error:", message=e)
        print(e)
        return None


def calculate_expire_time(given_time):
    """ This function takes the time in 100nanoseconds from active and convert it to days"""
    fetch_pasword_expiration = f'''
     $expiryTime = {given_time}
     $daysRemaining = [math]::ceiling(($expiryTime - (Get-Date).ToFileTime()) / 864000000000)
     $daysRemaining
     '''
    calculate_result = run_powershell_command(fetch_pasword_expiration, False)
    return int(calculate_result)


def search_by_employee_id(id):
    employee_id = id
    fetch_user_by_employeeid = f'''
    $employee_id = "{employee_id}"
    Get-ADUser -Filter "EmployeeID -eq $employee_id" -Properties sAMAccountName,EmployeeId,DisplayName,Mail,Manager,Title,physicalDeliveryOfficeName,Department,Enabled,"msDS-UserPasswordExpiryTimeComputed" |
    Select-Object sAMAccountName, EmployeeId, DisplayName, Mail, Manager, Title,physicalDeliveryOfficeName,Department,Enabled,"msDS-UserPasswordExpiryTimeComputed" |
    ConvertTo-Json
    '''
    search_result = run_powershell_command(fetch_user_by_employeeid)
    return search_result


def search_by_samaccountname(username):
    samaccount_name = username
    fetch_user_by_samaccountname = f'''
    $samaccountname = "{samaccount_name}"
    Get-ADUser -Filter {{sAMAccountName -eq $samaccountname}} -Properties sAMAccountName,EmployeeId,DisplayName,Mail,Title,Manager,physicalDeliveryOfficeName,Department,Enabled,"msDS-UserPasswordExpiryTimeComputed" |
    Select-Object sAMAccountName, EmployeeId, DisplayName, Mail, Title, Manager, physicalDeliveryOfficeName,Department,Enabled,"msDS-UserPasswordExpiryTimeComputed" |
    ConvertTo-Json
    '''
    search_result = run_powershell_command(fetch_user_by_samaccountname)
    return search_result


def clear_table():
    new_password_entry.delete(0, END)
    display_name_value.config(text="")
    email_value.config(text="")
    samaccountname_value.config(text="")
    job_title_value.config(text="")
    report_to_value.config(text="")
    department_value.config(text="")
    office_value.config(text="")
    user_status_value.config(text="")
    password_expires_at_value.config(text="")
    employee_id_value.config(text="")


def fill_the_table(search_result, password):
    """ Fills all the attributes and return the generated password"""
    sam_account_name = search_result.get("sAMAccountName", "")
    display_name = search_result.get("DisplayName", "")
    employeeid = search_result.get("EmployeeId", "")
    mail = search_result.get("Mail", "")
    title = search_result.get("Title", "")
    manager = search_result.get("Manager", "").split(",")[0].replace("CN=", "")
    department = search_result.get("Department", "")
    office = search_result.get("physicalDeliveryOfficeName", "")
    enabled = search_result.get("Enabled", "")
    status = lambda enabled_in_ad: "Active" if enabled_in_ad else "Inactive"
    password_expiry_time = calculate_expire_time(search_result.get("msDS-UserPasswordExpiryTimeComputed", ""))

    display_name_value.config(text=display_name)
    email_value.config(text=mail)
    samaccountname_value.config(text=sam_account_name)
    job_title_value.config(text=title)
    report_to_value.config(text=manager)
    department_value.config(text=department)
    office_value.config(text=office)
    user_status_value.config(text=status(enabled), font=("Arial", 12, "bold"))
    password_expires_at_value.config(text=f"{password_expiry_time} days")
    employee_id_value.config(text=employeeid)
    new_password_entry.insert(0, password)

def search(event):
    # event listens when hit Enter on keyboard to run search
    user_input = username_id_entry.get()
    clear_table()
    if len(user_input) != 0:
        if user_input.isnumeric():
            search_result = search_by_employee_id(user_input)
            if search_result:
                password = generate_password()
                fill_the_table(search_result, password)
                reset_password(search_result, password)
                return search_result
            else:
                messagebox.showerror(title="404", message=f"The user with {user_input} code not found.")

        elif isinstance(user_input, str):
            search_result = search_by_samaccountname(user_input)
            if search_result:
                password = generate_password()
                fill_the_table(search_result, password)
                reset_password(search_result, password)
                return search_result
            else:
                messagebox.showerror(title="404", message=f"The user  {user_input}  not found.")
    else:
        messagebox.showerror(title="OOPS!", message="This field should not be empty.")


def reset_password(user_object, password):
    pass


def send_to_printer(input):
    pass


def generate_password():
    # password_entry.delete(0, END)
    symbols = '!#$%@'
    numbers = '123456789'
    capital_letters = 'ABCDEFGHJKLMNOPQRSTUVWXYZ'
    lowercase_letters = 'abcdefghijkmnpqrstuvwxyz'

    # new_list = [new_item for item in list]
    password_capital_letters = [random.choice(capital_letters) for _ in range(3)]
    password_symbols = [random.choice(symbols) for _ in range(1)]
    password_numbers = [random.choice(numbers) for _ in range(4)]
    password_lowercase_letters = [random.choice(lowercase_letters) for _ in range(3)]

    password_list = password_capital_letters + password_numbers + password_symbols + password_lowercase_letters
    random.shuffle(password_list)
    password = ''.join(password_list)
    return password
    # password_entry.insert(0, password)
    # pyperclip.copy(password)

def email_correction(input):
    pass


def save_to_csv(input):
    pass


# UI-----------------------------------------------------UI
# UI-----------------------------------------------------UI
window = Tk()
window.title("Active Tools")
window.minsize(width=452, height=750)
window.config(bg=DEEP_BLUE)

# Search Section Frame
search_section_frame = Frame(window, bg=DEEP_BLUE, bd=2, relief=GROOVE)
search_section_frame.grid(row=0, column=0, columnspan=4, padx=10, pady=5, sticky="ew")

username_id_entry = Entry(search_section_frame, width=72)
username_id_entry.bind('<Return>', search)
username_id_entry.focus()
username_id_entry.config(highlightthickness=2, highlightcolor=STEEL_GREY, justify="left")
username_id_entry.grid(row=0, column=0, columnspan=4, pady=5, padx=5)

search_btn = Button(search_section_frame, text="Search", width=29, bg=CHARCOAL, fg="white", command=search)
search_btn.config(font=("Arial", 9, "bold"))
search_btn.grid(row=1, column=0, columnspan=2, pady=5, padx=2)

reset_and_print_btn = Button(search_section_frame, text="Reset Password And Print", width=30, bg=ACCENT_ORANGE,
                             fg="black")
reset_and_print_btn.config(font=("Arial", 9, "bold"), cursor="pirate")
reset_and_print_btn.grid(row=1, column=2, columnspan=2, pady=5)

# User Attribute Frame
user_attr_frame = Frame(window, bg=DEEP_BLUE, bd=2, relief=GROOVE)
user_attr_frame.grid(row=2, column=0, columnspan=4, padx=10, pady=5, sticky="ew")

# DisplayName
display_name_label = Label(user_attr_frame, text="Display Name: ")
display_name_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
display_name_label.grid(row=0, column=0, sticky="w")

display_name_value = Label(user_attr_frame, text="")
display_name_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
display_name_value.grid(row=0, column=2, sticky="w", padx=15)

# Email
email_label = Label(user_attr_frame, text="Email:")
email_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
email_label.grid(row=1, column=0, sticky="w")

email_value = Label(user_attr_frame, text="")
email_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
email_value.grid(row=1, column=2, sticky="w", padx=15)

# sAMAccountName
samaccountname_label = Label(user_attr_frame, text="SamAccountName:")
samaccountname_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
samaccountname_label.grid(row=2, column=0, sticky="w")

samaccountname_value = Label(user_attr_frame, text="")
samaccountname_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
samaccountname_value.grid(row=2, column=2, sticky="w", padx=15)

# Title
job_title_label = Label(user_attr_frame, text="Job Title:")
job_title_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
job_title_label.grid(row=3, column=0, sticky="w")

job_title_value = Label(user_attr_frame, text="")
job_title_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
job_title_value.grid(row=3, column=2, sticky="w", padx=15)

# Manager
report_to_label = Label(user_attr_frame, text="Report To:")
report_to_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
report_to_label.grid(row=4, column=0, sticky="w")

report_to_value = Label(user_attr_frame, text="")
report_to_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
report_to_value.grid(row=4, column=2, sticky="w", padx=15)

# Department
department_label = Label(user_attr_frame, text="Department:")
department_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
department_label.grid(row=5, column=0, sticky="w")

department_value = Label(user_attr_frame, text="")
department_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
department_value.grid(row=5, column=2, sticky="w", padx=15)

# physicalDeliveryOfficeName
office_label = Label(user_attr_frame, text="Office:")
office_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
office_label.grid(row=6, column=0, sticky="w")

office_value = Label(user_attr_frame, text="")
office_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
office_value.grid(row=6, column=2, sticky="w", padx=15)

# User Status (userAccountControl is 512 or 514 and 66050
user_status_label = Label(user_attr_frame, text="Status:")
user_status_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
user_status_label.grid(row=7, column=0, sticky="w")

user_status_value = Label(user_attr_frame, text="")
user_status_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
user_status_value.grid(row=7, column=2, sticky="w", padx=15)

# Time in days to password expiration (msDS-UserPasswordExpiryTimeComputed)
password_expires_at_label = Label(user_attr_frame, text="Password is valid until:")
password_expires_at_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
password_expires_at_label.grid(row=10, column=0, sticky="w")

password_expires_at_value = Label(user_attr_frame, text="")
password_expires_at_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12))
password_expires_at_value.grid(row=10, column=2, sticky="w", padx=15)

# EmployeeID
employee_id_label = Label(user_attr_frame, text="Employee ID:")
employee_id_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
employee_id_label.grid(row=11, column=0, sticky="w")

employee_id_value = Label(user_attr_frame, text="")
employee_id_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
employee_id_value.grid(row=11, column=2, sticky="w", padx=15)

# New generated Password
new_password = Label(user_attr_frame, text="New generated password: ")
new_password.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
new_password.grid(row=12, column=0, sticky="w")

new_password_entry = Entry(user_attr_frame, width=35)
new_password_entry.config(highlightthickness=2, highlightcolor=STEEL_GREY, justify="left")
new_password_entry.grid(row=12, column=2, columnspan=2, pady=5, padx=(15, 5))

# --------------------------------------------------------------
separator_canvas = Canvas(window, height=2, bg=CHARCOAL, highlightbackground=STEEL_GREY)
separator_canvas.grid(row=3, column=0, columnspan=4, sticky="nsew", pady=5, padx=10)
# The line inside the canvas that we use as separator
separator_canvas.create_line(20, 1, 432, 1, fill="white", width=10, dash=True)
# --------------------------------------------------------------

# Group Password Reset Frame
group_reset_title_frame = Frame(window, bd=2, relief=GROOVE, bg=DEEP_BLUE)
group_reset_title_frame.grid(row=4, column=0, columnspan=4, pady=5)

# Group password reset Title label
group_password_reset_label = Label(group_reset_title_frame, text="Group Password Reset", width=43, bg=DEEP_BLUE,
                                   fg="white")
group_password_reset_label.config(font=("Arial", 12, "bold"))
group_password_reset_label.grid(row=0, column=0, padx=5, pady=5)

# Frame for Email Textbox
email_textbox_frame = Frame(window, bd=2, relief=GROOVE, bg=DEEP_BLUE)
email_textbox_frame.grid(row=5, column=0, columnspan=4, padx=5, pady=(0, 10))

# Group Email textbox
group_email_textbox = Text(email_textbox_frame, height=15, width=55)
group_email_textbox.config(bg=STEEL_GREY, fg=DEEP_BLUE, font=("Arial", 11, "bold"))
group_email_textbox.grid(row=0, column=0, padx=3, pady=5)

# reset password and save to csv byn
reset_and_savecsv_btn = Button(email_textbox_frame, text="Reset and save to CSV")
reset_and_savecsv_btn.config(pady=5, width=62, bg=CHARCOAL, fg="white", cursor="pirate")
reset_and_savecsv_btn.grid(row=1, column=0, columnspan=4, pady=(0, 3))

window.mainloop()
