import tkinter
from tkinter import ttk
import re
import sqlite3
from tkinter import messagebox
import os
import openpyxl
import matplotlib.pyplot as plt
from collections import Counter

def enter_data():
    accepted = accept_var.get()
    
    if accepted=="Accepted":
        # User info
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()
        email = email_id_entry.get()

         # Email validation
        email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
 
        if firstname and lastname and re.match(email_regex, email):
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()
            gender = gender_var.get()
            
            
            # Course info
            registration_status = reg_status_var.get()
            numcourses = numcourses_spinbox.get()
            numsemesters = numsemesters_spinbox.get()
            
            print("First name: ", firstname, "Last name: ", lastname)
            print("Title: ", title, "Age: ", age, "Nationality: ", nationality, "Email:", email, "Gender:", gender)
            print("# Courses: ", numcourses, "# Semesters: ", numsemesters)
            print("Registration status", registration_status)
            print("------------------------------------------")

            # Create Table
            conn = sqlite3.connect('data.db')
            table_create_query = '''CREATE TABLE IF NOT EXISTS Student_Data 
                    (firstname TEXT, lastname TEXT, title TEXT, age INT, nationality TEXT, 
                    registration_status TEXT, num_courses INT, num_semesters INT)
            '''
            conn.execute(table_create_query)
            
            # Insert Data
            data_insert_query = '''INSERT INTO Student_Data (firstname, lastname, title, 
            age, nationality, registration_status, num_courses, num_semesters) VALUES 
            (?, ?, ?, ?, ?, ?, ?, ?)'''
            data_insert_tuple = (firstname, lastname, title,
                                  age, nationality, registration_status, numcourses, numsemesters)
            cursor = conn.cursor()
            cursor.execute(data_insert_query, data_insert_tuple)
            conn.commit()
            conn.close()

            
            filepath = "F:\Hope Foundation Course\python\data.xlsx"
            
            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["Title","First Name", "Last Name","Gender", "Age", "Nationality", "Email",
                           "# Courses", "# Semesters", "Registration status"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([title , firstname, lastname, gender, age, nationality,email, numcourses,
                          numsemesters, registration_status])
            workbook.save(filepath)
                
        else:
            tkinter.messagebox.showwarning(title="Error", message="First name and last name are required. Invalid email")
    else:
        tkinter.messagebox.showwarning(title= "Error", message="You have not accepted the terms")

def show_charts():
    plt.rcParams["axes.prop_cycle"] = plt.cycler(
    color=["#4C2A85", "#BE96FF", "#957DAD", "#5E366E", "#A98CCC"])

    conn = sqlite3.connect('data.db')
    cursor = conn.cursor()

  # Fetch Age & Nationality data
    cursor.execute("SELECT nationality, age FROM Student_Data")
    data = cursor.fetchall()

    # Fetch Nationality counts
    cursor.execute("SELECT nationality FROM Student_Data")
    nationality_list = cursor.fetchall()  # Fetch data properly
    conn.close()
     
     

    if not data:
        messagebox.showwarning("No Data", "No data found in the database.")
        return
    
    nationalities, ages = zip(*data)
    nationality_counts = Counter(n[0] for n in nationality_list)
    #registration_counts = Counter(registration_list)

     # Plot first bar chart (Age by Nationality)
    plt.figure(figsize=(12, 5))
    plt.subplot(1, 2, 1)
    plt.bar(nationalities, ages)
    plt.xlabel("Nationality")
    plt.ylabel("Age")
    plt.title("Age by Nationality")
     
    plt.ylim(17, max(ages) + 2)

     # Plot second bar chart (Students per Nationality)
    plt.subplot(1, 2, 2)
    plt.pie(nationality_counts.values(), labels=nationality_counts.keys(), autopct='%1.1f%%')

    plt.title("Students per Nationality")
    
    plt.show()

 

window = tkinter.Tk()
window.title("Data Entry Form")

frame = tkinter.Frame(window)
frame.pack()

# Saving User Info
user_info_frame =tkinter.LabelFrame(frame, text="User Information")
user_info_frame.grid(row= 0, column=0, padx=20, pady=10)

title_label = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=["", "Mr.", "Ms.", "Dr."])
title_label.grid(row=0, column=0)
title_combobox.grid(row=1, column=0)



first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row=0, column=1)
last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0, column=2)

first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=1)
last_name_entry.grid(row=1, column=2)



age_label = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

nationality_label = tkinter.Label(user_info_frame, text="Nationality")
nationality_combobox = ttk.Combobox(user_info_frame, values=["India", "Africa", "Antarctica", "Asia", "Europe", "North America", "Oceania", "South America"])
nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)

email_id_label = tkinter.Label(user_info_frame, text="Email.")
email_id_label .grid(row=2, column=2)
email_id_entry = tkinter.Entry(user_info_frame)
email_id_entry.grid(row=3, column=2)


# Gender Label
gender_label = tkinter.Label(user_info_frame, text="Gender")
gender_label.grid(row=4, column=0)

# Gender Selection (Radio Buttons)
gender_var = tkinter.StringVar(value="None")
male_radiobutton = tkinter.Radiobutton(user_info_frame, text="Male", variable=gender_var, value="Male")
female_radiobutton = tkinter.Radiobutton(user_info_frame, text="Female", variable=gender_var, value="Female")
other_radiobutton = tkinter.Radiobutton(user_info_frame, text="Other", variable=gender_var, value="Other")

male_radiobutton.grid(row=5, column=0)
female_radiobutton.grid(row=5, column=1)
other_radiobutton.grid(row=5, column=2)


for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)



# Saving Course Info
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

registered_label = tkinter.Label(courses_frame, text="Registration Status")

reg_status_var = tkinter.StringVar(value="Not Registered")
registered_check = tkinter.Checkbutton(courses_frame, text="Currently Registered",
                                       variable=reg_status_var, onvalue="Registered", offvalue="Not registered")

registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

numcourses_label = tkinter.Label(courses_frame, text= "# Completed Courses")
numcourses_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)

numsemesters_label = tkinter.Label(courses_frame, text="# Semesters")
numsemesters_spinbox = tkinter.Spinbox(courses_frame, from_=0, to="infinity")
numsemesters_label.grid(row=0, column=2)
numsemesters_spinbox.grid(row=1, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text= "I accept the terms and conditions.",
                                  variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

# Button
button = tkinter.Button(frame, text="Enter data", command= enter_data,  bg="green", fg="white",font=("Arial", 10, "bold"))
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

chart_button = tkinter.Button(frame, text="Show Charts", command=show_charts, bg="orange", fg="white",font=("Arial", 10, "bold"))
chart_button.grid(row=4, column=0, padx=20, pady=10)
 
window.mainloop()