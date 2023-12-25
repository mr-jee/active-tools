from tkinter import *
from tkinter import messagebox
import random

DEEP_BLUE = "#004080"
ACCENT_ORANGE = "#ffa500"
STEEL_GREY = "#808080"
CHARCOAL = "#333333"





# UI-----------------------------------------------------UI
# UI-----------------------------------------------------UI
window = Tk()
window.title("Active Tools")
window.minsize(width=450, height=700)
window.config(bg=DEEP_BLUE)

username_id_entry = Entry(width=72)
username_id_entry.focus()
username_id_entry.grid(row=0,column=0,columnspan=4, pady=5, padx=5)

search_btn = Button(text="Search", width=29)
search_btn.grid(row=1, column=0, columnspan=2, pady=5, padx=5)

reset_and_print_btn = Button(text="Reset Password And Print", width=29)
reset_and_print_btn.grid(row=1, column=2, columnspan=2, pady=5)

display_name_label = Label(text="Display Name: ")
display_name_label.grid(row=2, column=0, sticky="w", padx=5)

display_name_value = Label(text="***")
display_name_value.grid(row=3, column=0, sticky="w", padx=5)

email_label = Label(text="Email:")
email_label.grid(row=2, column=2, sticky="w", padx=5)

email_value = Label(text="***@testi.local")
email_value.grid(row=3, column=2, sticky="w", padx=5)

samaccountname_label = Label(text="SamAccountName:")
samaccountname_label.grid(row=4, column=0, sticky="w", padx=5)

samaccountname_value = Label(text="***")
samaccountname_value.grid(row=5, column=0, sticky="w", padx=5)




window.mainloop()
