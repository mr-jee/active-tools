from tkinter import *
from tkinter import messagebox
import random
import subprocess
import json

DEEP_BLUE = "#004080"
ACCENT_ORANGE = "#ffa500"
STEEL_GREY = "#9a9a9a"
CHARCOAL = "#333333"

# Back-----------------------------------------------------Back
# I had some help from chatGPT for this function, so I wrote complete help wherever I got it from chat
def run_powershell_command(powershell_script, expect_json=True):
    ''' This method meant to run powershell commands, it takes powershell command as a variable, if JSON output is needed(like data of a user) expect_json
    equals True, if no json needed, set it to False'''
    try:
        result = subprocess.run(["powershell", "-Command", powershell_script], capture_output=True, text=True, check=True)

        if expect_json:
            try:
                # Try parsing the output as JSON
                # stdout is standard output so by this feature we can detect the exact error that is happened
                messagebox.showinfo(title="Runs OK", message=result.stdout)
                return json.loads(result.stdout)
            # If there is a problem in the JSON output this line will show us the error
            except json.JSONDecodeError:
                # Return None if JSON parsing fails
                messagebox.showinfo(title="JSONDecode ERROR", message=result.stdout)
                print("Error decoding JSON. Output:", result.stdout)
                return None
        else:
            # If the failure is not bcs of JSON it shows the main stdout
            # Return the raw text output
            return result.stdout
    # When none of the above happens  I expect it would be the text of powershell error which is related to check=True in the arguments
    # When you use subprocess.run() to execute an external command, it runs the command in a separate process.
    # If the command completes successfully, it returns with a zero exit code.
    # If the command encounters an error or fails, it returns with a non-zero exit code.
    except subprocess.CalledProcessError as e:
        # Handle PowerShell errors
        messagebox.showerror(title="PowerShell Error:", message=e)
        print("PowerShell Error:", e)
        return None




















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
username_id_entry.focus()
username_id_entry.config(highlightthickness=2, highlightcolor=STEEL_GREY, justify="left")
username_id_entry.grid(row=0, column=0, columnspan=4, pady=5, padx=5)

search_btn = Button(search_section_frame, text="Search", width=29, bg=CHARCOAL, fg="white")
search_btn.config(font=("Arial", 9, "bold"))
search_btn.grid(row=1, column=0, columnspan=2, pady=5, padx=2)

reset_and_print_btn = Button(search_section_frame, text="Reset Password And Print", width=30, bg=ACCENT_ORANGE, fg="black")
reset_and_print_btn.config(font=("Arial", 9, "bold"), cursor="pirate")
reset_and_print_btn.grid(row=1, column=2, columnspan=2, pady=5)


# User Attribute Frame
user_attr_frame = Frame(window, bg=DEEP_BLUE, bd=2, relief=GROOVE)
user_attr_frame.grid(row=2, column=0, columnspan=4, padx=10, pady=5, sticky="ew")

# DisplayName
display_name_label = Label(user_attr_frame, text="Display Name: ")
display_name_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
display_name_label.grid(row=0, column=0, sticky="w")

display_name_value = Label(user_attr_frame, text="***")
display_name_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
display_name_value.grid(row=0, column=2, sticky="w", padx=15)

# Email
email_label = Label(user_attr_frame, text="Email:")
email_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
email_label.grid(row=1, column=0, sticky="w")

email_value = Label(user_attr_frame, text="***@testi.local")
email_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
email_value.grid(row=1, column=2, sticky="w", padx=15)

# sAMAccountName
samaccountname_label = Label(user_attr_frame, text="SamAccountName:")
samaccountname_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
samaccountname_label.grid(row=2, column=0, sticky="w")

samaccountname_value = Label(user_attr_frame, text="***")
samaccountname_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
samaccountname_value.grid(row=2, column=2, sticky="w", padx=15)

# Title
job_title_label = Label(user_attr_frame, text="Job Title:")
job_title_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
job_title_label.grid(row=3, column=0, sticky="w")

job_title_value = Label(user_attr_frame, text="***")
job_title_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
job_title_value.grid(row=3, column=2, sticky="w", padx=15)

# Manager
report_to_label = Label(user_attr_frame, text="Report To:")
report_to_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
report_to_label.grid(row=4, column=0, sticky="w")

report_to_value = Label(user_attr_frame, text="***")
report_to_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
report_to_value.grid(row=4, column=2, sticky="w", padx=15)

# Department
department_label = Label(user_attr_frame, text="Department:")
department_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
department_label.grid(row=5, column=0, sticky="w")

department_value = Label(user_attr_frame, text="***")
department_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
department_value.grid(row=5, column=2, sticky="w", padx=15)

# physicalDeliveryOfficeName
office_label = Label(user_attr_frame, text="Office:")
office_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
office_label.grid(row=6, column=0, sticky="w")

office_value = Label(user_attr_frame, text="***")
office_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
office_value.grid(row=6, column=2, sticky="w", padx=15)

# User Status (userAccountControl is 512 or 514 and 66050
user_status_label = Label(user_attr_frame, text="Status:")
user_status_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
user_status_label.grid(row=7, column=0, sticky="w")

user_status_value = Label(user_attr_frame, text="***")
user_status_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 10))
user_status_value.grid(row=7, column=2, sticky="w", padx=15)

# Time in days to password expiration (msDS-UserPasswordExpiryTimeComputed)
password_expires_at_label = Label(user_attr_frame, text="Password is valid until:")
password_expires_at_label.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
password_expires_at_label.grid(row=10, column=0, sticky="w")

password_expires_at_value = Label(user_attr_frame, text="*** days")
password_expires_at_value.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
password_expires_at_value.grid(row=10, column=2, sticky="w", padx=15)

# New generated Password
new_password = Label(user_attr_frame, text="New generated password: ")
new_password.config(bg=DEEP_BLUE, fg="white", font=("Arial", 12, "bold"))
new_password.grid(row=11, column=0, sticky="w")

new_password_entry = Entry(user_attr_frame, width=35)
new_password_entry.config(highlightthickness=2, highlightcolor=STEEL_GREY, justify="left")
new_password_entry.grid(row=11, column=2, columnspan=2, pady=5, padx=(15,5))

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
group_password_reset_label = Label(group_reset_title_frame, text="Group Password Reset", width=43, bg=DEEP_BLUE, fg="white")
group_password_reset_label.config(font=("Arial", 12, "bold"))
group_password_reset_label.grid(row=0, column=0, padx=5, pady=5)

# Frame for Email Textbox
email_textbox_frame = Frame(window, bd=2, relief=GROOVE, bg=DEEP_BLUE)
email_textbox_frame.grid(row=5, column=0, columnspan=4, padx=5)

# Group Email textbox
group_email_textbox = Text(email_textbox_frame, height=15, width=55)
group_email_textbox.config(bg=STEEL_GREY, fg=DEEP_BLUE, font=("Arial", 11, "bold"))
group_email_textbox.grid(row=0, column=0, padx=3, pady=5)

# reset password and save to csv byn
reset_and_savecsv_btn = Button(email_textbox_frame, text="Reset and save to CSV")
reset_and_savecsv_btn.config(pady=5, width=62, bg=CHARCOAL, fg="white", cursor="pirate")
reset_and_savecsv_btn.grid(row=1,column=0, columnspan=4)

window.mainloop()
