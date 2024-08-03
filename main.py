import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchDriverException, ElementNotInteractableException
import sys
import os
import sqlite3
import pandas as pd
import threading
import time
import ctypes
from credentials import user_name, password, url

class DataImportApp:
    def __init__(self, root):
        # Defining Base Path
        if hasattr(sys, '_MEIPASS'):
            self.base_path = sys._MEIPASS
        else:
            self.base_path = os.path.abspath(".")
        # Creating Root Window
        self.root = root
        self.root.title("Employee Data Import App")
        self.screen_height = self.root.winfo_screenheight()
        self.screen_width = self.root.winfo_screenwidth()
        # Defining Taskbar Icon
        taskbar_icon = 'mycompany.myproduct.subproduct.version'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(taskbar_icon)
        self.root.iconbitmap(os.path.join(self.base_path, "import.ico"))
        # Adjusting Root Window's Appearance
        self.root.config(background="#caf4fe")
        self.root.attributes("-alpha","0.95")
        self.root.geometry(f"320x150+{(self.screen_width - 320) // 2}+{(self.screen_height - 150) // 2}")
        self.root.resizable(False, False)
        # Description Label
        self.description = ttk.Label(self.root, text="Please upload excel file, then press start:",background="#caf4fe",foreground="black", font=("Arial", 10))
        self.description.grid(row=0, column=0, columnspan=6, sticky="w", pady=10)
        # Blank space reserved for progress bar
        self.progress_bar_frame = ttk.Label(self.root, text="", background="#caf4fe")
        self.progress_bar_frame.grid(row=1, column=0, columnspan=6, pady=10)
        # Upload status label
        self.status_label = ttk.Label(self.root, text="", background="#caf4fe", font=("Arial",8))
        self.status_label.grid(row=2, column=0, columnspan=6)
        # File selection button
        self.select_button = ttk.Button(self.root, text="Upload File", command=self.select_excel)
        self.select_button.grid(row=3, column=0, columnspan=2, padx=10, pady=10)
        # Start button
        self.start_button = ttk.Button(self.root, text="Start Process", command=self.start_process, state="disabled")
        self.start_button.grid(row=3, column=2, columnspan=2, padx=10, pady=10)
        # Quit button
        self.change_button = ttk.Button(self.root, text="Change Credentials", command=self.open_credentials_window)
        self.change_button.grid(row=3, column=4, columnspan=2, padx=10, pady=10)
        # Database Setup
        self.connector = sqlite3.connect("credentials.db")
        self.cursor = self.connector.cursor()
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS user_names (
                user_name text
            )
        """)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS passwords (
                password text
            )
        """)
        self.connector.commit()
        # User Name
        self.cursor.execute("""
            SELECT * FROM user_names;
        """)
        try:
            self.user_name = self.cursor.fetchall()[-1][0]
        except IndexError:
            self.user_name = user_name
        # Password
        self.cursor.execute("""
            SELECT * FROM passwords;
        """)
        try:
            self.password = self.cursor.fetchall()[-1][0]
        except IndexError:
            self.password = password
        self.mandatory_columns = ["code", "first_name", "father_name", "last_name", "first_name_fa", "father_name_fa", "last_name_fa", "national_number", "email", "birth_date", "gender", "marital_status", "site", "hierachy", "position", "join_date", "contract_type", "contract_start_date", "contract_end_date", "work_type", "work_class", "manager", "grade", "basic_salary", "phone_number", "place_of_issue", "birth_certificate_no", "birth_certificate_serial", "place_of_issue_fa", "address_fa", "position_fa", "hrbp"]

    def select_excel(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if self.file_path:
            self.upload_excel()
    
    def upload_excel(self):
        # Progress Bar
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
        self.progress_bar.grid(row=1, column=1, columnspan=4)
        self.progress_value = 0
        self.progress_bar["value"] = self.progress_value
        self.status_label.config(text=f"Uploading {self.file_path.split('/')[-1]}...")
        self.upload_thread = threading.Thread(target=self.upload_process)
        self.upload_thread.start()
        self.check_upload_progress()
    
    def upload_process(self):
        # Simulate uploading (replace with actual upload logic)
        total_size = 100  # Total size of the file (for simulation)
        while self.progress_value < total_size:
            time.sleep(0.1)  # Simulate delay
            # Update progress
            self.progress_value += 10  # Increase progress (for simulation)  
        # Once uploaded, read data from Excel
        self.df = pd.read_excel(self.file_path)
        self.df.drop(index=0, axis=0, inplace=True)
        missing_columns = []
        for column in self.mandatory_columns:
            if column not in self.df.columns:
                missing_columns.append(column)
        if missing_columns:
            self.start_button.config(state="disabled")
            messagebox.showerror("Missing Columns", f"The uploaded file is missing the following columns:\n{'\n'.join(missing_columns)}")
        else:
            self.start_button.config(state="normal")
        self.status_label.config(text=f"{self.file_path.split("/")[-1]} Uploaded")
    
    def check_upload_progress(self):
        if self.progress_value < 100:
            self.progress_bar["value"] = self.progress_value
            self.root.after(100, self.check_upload_progress)  # Check progress every 100ms
        else:
            self.progress_bar["value"] = 100
    
    def open_credentials_window(self):
        # Create Password Window
        self.credentials_window = tk.Toplevel(self.root)
        self.credentials_window.title("Change Credentials")
        self.credentials_window.iconbitmap(os.path.join(self.base_path, "password.ico"))
        self.credentials_window.config(background="#caf4fe")
        self.credentials_window.attributes("-alpha","0.95")
        self.credentials_window.geometry(f"320x90+{(self.screen_width - 320) // 2}+{(self.screen_height - 60) // 2}")
        self.credentials_window.resizable(False, False)
        # Credentials Window's Buttons
        self.change_user_name_button = ttk.Button(self.credentials_window, text="Change User Name", command=self.open_user_name_window)
        self.change_user_name_button.grid(row=0, column=0, padx=25, pady=30)
        self.change_password_button = ttk.Button(self.credentials_window, text="Change Password", command=self.open_password_window)
        self.change_password_button.grid(row=0, column=1, padx=25, pady=30)

    def open_user_name_window(self):
        # Create user_name Window
        self.user_name_window = tk.Toplevel(self.root)
        self.user_name_window.title("Change User Name")
        self.user_name_window.iconbitmap(os.path.join(self.base_path, "username.ico"))
        self.user_name_window.config(background="#caf4fe")
        self.user_name_window.attributes("-alpha","0.95")
        self.user_name_window.geometry(f"320x130+{(self.screen_width - 320) // 2}+{(self.screen_height - 90) // 2}")
        self.user_name_window.resizable(False, False)
        self.user_name_window.bind("<Return>", lambda event: self.change_user_name())
        # Current user_name Label and Entry
        self.current_user_name_label = ttk.Label(self.user_name_window, text="Current User Name: ", background="#caf4fe", font=("Arial",8))
        self.current_user_name_label.grid(row=0, column=0, sticky="w", pady=10)
        self.current_user_name_entry = ttk.Entry(self.user_name_window, width=34, font=("Arial", 8))
        self.current_user_name_entry.grid(row=0, column=1, columnspan=2, pady=10, sticky="e")
        # New user_name Label and Entry
        self.new_user_name_label = ttk.Label(self.user_name_window, text="New User Name: ", background="#caf4fe", font=("Arial",8))
        self.new_user_name_label.grid(row=1, column=0, sticky="w", pady=10)
        self.new_user_name_entry = ttk.Entry(self.user_name_window, width=34, font=("Arial", 8))
        self.new_user_name_entry.grid(row=1, column=1, columnspan=2, pady=10, sticky="e")
        # user_name Window's Buttons
        self.blank_label = ttk.Label(self.user_name_window, text="", background="#caf4fe")
        self.blank_label.grid(row=2, column=0, pady=10)
        self.change_user_name_button = ttk.Button(self.user_name_window, text="Change User Name", command=self.change_user_name)
        self.change_user_name_button.grid(row=2, column=1, pady=10, sticky="w")

    def change_user_name(self):
        if self.current_user_name_entry.get() != self.user_name:
            messagebox.showerror("Wrong User Name", "Current user name is incorrect.", parent=self.user_name_window)
        elif self.current_user_name_entry.get() == self.new_user_name_entry.get():
            messagebox.showerror("Repeated User Name", "User name is repeated.", parent=self.user_name_window)
        else:
            self.cursor.execute("""
                INSERT INTO user_names VALUES (?)
            """, (self.new_user_name_entry.get(),))
            self.connector.commit()
            self.cursor.execute("""
                SELECT * FROM user_names;
            """)
            self.user_name = self.cursor.fetchall()[-1][0]
            messagebox.showinfo("User Name Changed", "User name changed successfully.", parent=self.user_name_window)
            self.user_name_window.destroy()

    def open_password_window(self):
        # Create Password Window
        self.password_window = tk.Toplevel(self.root)
        self.password_window.title("Change Password")
        self.password_window.iconbitmap(os.path.join(self.base_path, "password.ico"))
        self.password_window.config(background="#caf4fe")
        self.password_window.attributes("-alpha","0.95")
        self.password_window.geometry(f"320x130+{(self.screen_width - 320) // 2}+{(self.screen_height - 90) // 2}")
        self.password_window.resizable(False, False)
        self.password_window.bind("<Return>", lambda event: self.change_password())
        # Current Password Label and Entry
        self.current_password_label = ttk.Label(self.password_window, text="Current Password: ", background="#caf4fe", font=("Arial",8))
        self.current_password_label.grid(row=0, column=0, sticky="w", pady=10)
        self.current_password_entry = ttk.Entry(self.password_window, width=35, font=("Arial", 8), show="*")
        self.current_password_entry.grid(row=0, column=1, pady=10, sticky="e")
        # New Password Label and Entry
        self.new_password_label = ttk.Label(self.password_window, text="New Password: ", background="#caf4fe", font=("Arial",8))
        self.new_password_label.grid(row=1, column=0, sticky="w", pady=10)
        self.new_password_entry = ttk.Entry(self.password_window, width=35, font=("Arial", 8), show="*")
        self.new_password_entry.grid(row=1, column=1, pady=10, sticky="e")
        # Password Window's Buttons
        self.show_password_button = ttk.Button(self.password_window, text="Show Passwords", command=self.show_passwords)
        self.show_password_button.grid(row=2, column=0, pady=10, sticky="e")
        self.change_password_button = ttk.Button(self.password_window, text="Change Password", command=self.change_password)
        self.change_password_button.grid(row=2, column=1, pady=10, sticky="e")

    def show_passwords(self):
        self.current_password_entry.config(show="")
        self.new_password_entry.config(show="")
        self.show_password_button.config(text="Hide Passwords", command=self.hide_passwords)
    
    def hide_passwords(self):
        self.current_password_entry.config(show="*")
        self.new_password_entry.config(show="*")
        self.show_password_button.config(text="Show Passwords", command=self.show_passwords)

    def change_password(self):
        if self.current_password_entry.get() != self.password:
            messagebox.showerror("Wrong Password", "Current password is incorrect.", parent=self.password_window)
        elif self.current_password_entry.get() == self.new_password_entry.get():
            messagebox.showerror("Repeated Password", "Password is repeated.", parent=self.password_window)
        else:
            self.cursor.execute("""
                INSERT INTO passwords VALUES (?)
            """, (self.new_password_entry.get(),))
            self.connector.commit()
            self.cursor.execute("""
                SELECT * FROM passwords;
            """)
            self.password = self.cursor.fetchall()[-1][0]
            messagebox.showinfo("Password Changed", "Password changed successfully.", parent=self.password_window)
            self.password_window.destroy()

    def start_process(self):
        try:
            service = Service(executable_path="chromedriver.exe")
            self.driver = webdriver.Chrome(service=service)
            self.wait = WebDriverWait(self.driver, 30)
        except NoSuchDriverException:
            messagebox.showerror("Driver Error", "Please add chromedriver.exe file in the directory.")
            sys.exit()
        try:
            # Open Website
            self.driver.get(url)
            # Maximize Window
            self.driver.maximize_window()
            # Login
            self.wait.until(
                EC.visibility_of_element_located((By.ID, 'btnLogin'))
                ).click()
            # Sign in
            self.wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "a[href='/Account/RequestAuthentication']"))
                ).click()
            # User_name
            self.wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Account']"))
                ).send_keys(self.user_name)
            # Password
            self.driver.find_element(By.CSS_SELECTOR, "input[placeholder='Password']").send_keys(self.password)
            # Submit
            self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
            # Configuring Stop Process Button
            self.start_button.config(text="Stop Process", command=self.stop_process)
            # PIN Code Entry Window
            self.open_pin_window()
        except:
            self.stop_process()
            messagebox.showerror("Network Error", "Network is unstable.")

    def open_pin_window(self):
        # Create pin Window
        self.pin_window = tk.Toplevel(self.root)
        self.pin_window.title("Authentication Code")
        self.pin_window.iconbitmap(os.path.join(self.base_path, "password.ico"))
        self.pin_window.config(background="#caf4fe")
        self.pin_window.attributes("-alpha","0.95")
        self.pin_window.geometry(f"320x130+{(self.screen_width - 320) // 2}+{(self.screen_height - 90) // 2}")
        self.pin_window.resizable(False, False)
        self.pin_window.bind("<Return>", lambda event: self.confirm_pin())
        # PIN Label and Entry
        self.current_pin_label = ttk.Label(self.pin_window, text="Please enter the code from Microsoft Authenticator: ", background="#caf4fe", font=("Arial", 10))
        self.current_pin_label.grid(row=0, column=0, sticky="w", pady=10)
        self.current_pin_entry = ttk.Entry(self.pin_window, width=40, font=("Arial", 8))
        self.current_pin_entry.grid(row=1, column=0, pady=10)
        # pin Window's Buttons
        self.confirm_button = ttk.Button(self.pin_window, text="Confirm", command=self.confirm_pin)
        self.confirm_button.grid(row=2, column=0, pady=10)
        self.current_pin_entry.focus()

    def confirm_pin(self):
        self.pin = self.current_pin_entry.get()
        self.pin_window.destroy()
        self.continue_process()
    
    def continue_process(self):
        try:
            # PIN_Code
            self.wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Enter PIN code']"))
                ).send_keys(self.pin)
            # Validate
            self.driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
            # People
            self.wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "a[href='/Employees/Dashboard/Index']"))
                ).click()
            # View Employees
            self.wait.until(
                EC.visibility_of_element_located((By.CSS_SELECTOR, "a[href='/Employees/Home/Index']"))
                ).click()
            # Iterating through Rows
            for _, row in self.df.iterrows():
                try:
                    self.driver.find_element(By.CSS_SELECTOR, "button[aria-controls='liactions']").click()
                except ElementNotInteractableException:
                    pass
                self.wait.until(
                    EC.visibility_of_element_located((By.ID, 'btn_Add'))
                    ).click()
                # code
                code = row["code"]
                code_entry = self.wait.until(
                    EC.visibility_of_element_located((By.NAME, 'Code'))
                    )
                code_entry.clear()
                code_entry.send_keys(code)
                # first_name
                first_name = row["first_name"]
                self.driver.find_element(By.NAME, 'lstNames[0].First').send_keys(first_name)
                # Additonal Fields
                self.driver.find_element(By.ID, 'OtherLanguageEmployeeNames_btn').click()
                # father_name
                father_name = row["father_name"]
                self.wait.until(
                    EC.visibility_of_element_located((By.NAME, 'lstNames[0].Second'))
                    ).send_keys(father_name)
                # last_name
                last_name = row["last_name"]
                self.driver.find_element(By.NAME, 'lstNames[0].Last').send_keys(last_name)
                # first_name_fa
                first_name_fa = row["first_name_fa"]
                self.driver.find_element(By.NAME, 'lstNames[1].First').send_keys(first_name_fa)
                # father_name_fa
                father_name_fa = row["father_name_fa"]
                self.driver.find_element(By.NAME, 'lstNames[1].Second').send_keys(father_name_fa)
                # last_name_fa
                last_name_fa = row["last_name_fa"]
                self.driver.find_element(By.NAME, 'lstNames[1].Last').send_keys(last_name_fa)
                # national_number
                national_number = row["national_number"]
                self.driver.find_element(By.NAME, 'NationalNumber').send_keys(national_number)
                # email
                email = row["email"]
                self.driver.find_element(By.NAME, 'Email').send_keys(email)
                # Date Pickers
                date_pickers = self.driver.find_elements(By.CLASS_NAME, "pwt-datepicker-input-element")
                # birth_date
                birth_date = row["birth_date"]
                birth_date_entry = date_pickers[2]
                birth_date_entry.send_keys(birth_date)
                birth_date_entry.send_keys(Keys.TAB)
                # Drop Downs
                drop_downs = self.driver.find_elements(By.CLASS_NAME, "select2-selection__rendered")
                # gender
                gender = row["gender"]
                genders = {"مرد": "Male", "زن": "Female"}
                drop_downs[1].click()
                for gender_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if gender_.text == genders[gender]:
                        gender_.click()
                        break
                # marital_status
                marital_status = row["marital_status"]
                marital_status_dict = {"مجرد": "Single", "متاهل": "Married" ,"متارکه": "Divorced" , "بیوه": "Widowed"}
                drop_downs[2].click()
                for marital_status_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if marital_status_.text == marital_status_dict[marital_status]:
                        marital_status_.click()
                        break
                # religion
                drop_downs[3].click()
                for religion_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if religion_.text == "Other":
                        religion_.click()
                        break
                # nationality
                drop_downs[4].click()
                for nationality_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if nationality_.text == "Iranian":
                        nationality_.click()
                        break
                # Second Tab
                self.driver.find_elements(By.CLASS_NAME, "BT-line-tabs-link")[-1].click()
                # site
                site = row["site"]
                drop_downs[7].click()
                for site_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if site_.text == site:
                        site_.click()
                        break
                # location
                drop_downs[8].click()
                for location_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if location_.text == site:
                        location_.click()
                        break
                # Tree Selections
                tree_selectors = self.driver.find_elements(By.CLASS_NAME, "treeselect-input__edit")
                # hierachy
                hierachy = row["hierachy"]
                hierachies = {"SC": 2, "Fin": 3, "IT": 4, "HR": 7, "CORA": 8}
                hierachy_entry = tree_selectors[1]
                hierachy_entry.click()
                hierachy_entry.send_keys(Keys.ARROW_RIGHT)
                hierachy_entry.send_keys(Keys.ARROW_DOWN)
                hierachy_entry.send_keys(Keys.ARROW_RIGHT)
                if hierachy in hierachies:
                    for _ in range(hierachies[hierachy]):
                        hierachy_entry.send_keys(Keys.ARROW_DOWN)
                hierachy_entry.send_keys(Keys.RETURN)
                time.sleep(1)
                # position
                position = row["position"]
                drop_downs[9].click()
                self.driver.find_element(By.CLASS_NAME, "select2-search__field").send_keys(position)
                for position_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if position_.text == position:
                        position_.click()
                        break
                # hiring_date
                join_date = row["join_date"]
                hiring_date_entry = date_pickers[3]
                hiring_date_entry.send_keys(join_date)
                hiring_date_entry.send_keys(Keys.TAB)
                # join_date
                join_date_entry = date_pickers[4]
                join_date_entry.send_keys(join_date)
                join_date_entry.send_keys(Keys.TAB)
                # contract_type
                contract_type = row["contract_type"]
                drop_downs[12].click()
                self.driver.find_element(By.CLASS_NAME, "select2-search__field").send_keys(contract_type)
                for contract_type_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if contract_type_.text == contract_type:
                        contract_type_.click()
                        break
                # contract_start_date
                contract_start_date = row["contract_start_date"]
                contract_start_date_entry = date_pickers[5]
                # contract_start_date_entry = self.wait.until(
                #     EC.visibility_of_element_located((By.XPATH, f"/html/body/div[{8 + 7*(i-1)}]/div/div/div/form/div/div/div[2]/div/div[1]/div[1]/div/div[9]/div[2]/div[2]/div[1]/div/div/div/div/input"))
                # )
                contract_start_date_entry.send_keys(contract_start_date)
                contract_start_date_entry.send_keys(Keys.TAB)
                # contract_end_date
                contract_end_date = row["contract_end_date"]
                contract_end_date_entry = date_pickers[6]
                contract_end_date_entry.send_keys(contract_end_date)
                contract_end_date_entry.send_keys(Keys.TAB)
                # work_type
                work_type = row["work_type"]
                drop_downs[14].click()
                self.driver.find_element(By.CLASS_NAME, "select2-search__field").send_keys(work_type)
                for work_type_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if work_type_.text == work_type:
                        work_type_.click()
                        break
                # work_class
                work_class = row["work_class"]
                drop_downs[15].click()
                self.driver.find_element(By.CLASS_NAME, "select2-search__field").send_keys(work_class)
                for work_class_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if work_class_.text == work_class:
                        work_class_.click()
                        break
                # managers
                manager = row["manager"]
                managers = manager.split(", ")
                self.driver.find_element(By.CLASS_NAME, "btn-employeePicker-find").click()
                search_bar = self.wait.until(
                    EC.visibility_of_element_located((By.ID, "txtSearch_modal"))
                )
                for manager_ in managers:
                    search_bar.clear()
                    search_bar.send_keys(manager_)
                    search_bar.send_keys(Keys.RETURN)
                    time.sleep(1)
                    self.driver.find_elements(By.CLASS_NAME, "BT_employee_search_selection")[0].click()
                self.wait.until(
                    EC.visibility_of_element_located((By.ID, "btnchoose_submit"))
                ).click()
                # grade
                grade = row["grade"]
                grades_1 = ["WL1A", "WL1B", "WL1C", "WL1D", "UFLP", "WL2A", "WL2B", "WL2C", "WL3X", "WL4X", "F1D", "F1E", "F1F", "F2D", "F2E", "F2F", "F3D", "F3E", "F3F"]
                grades_2 = {"F4D": 20, "ULIP": 21, "F4E": 25, "F4F": 26, "F0": 28, "Contractor": 30}
                grade_entry = tree_selectors[2]
                grade_entry.click()
                grade_entry.send_keys(Keys.ARROW_RIGHT)
                grade_entry.send_keys(Keys.ARROW_DOWN)
                if grade in grades_1:
                    for _ in range(grades_1.index(grade)):
                        grade_entry.send_keys(Keys.ARROW_DOWN)
                elif grade in grades_2:
                    for _ in range(grades_2[grade]):
                        grade_entry.send_keys(Keys.ARROW_DOWN)
                grade_entry.send_keys(Keys.ARROW_RIGHT)
                grade_entry.send_keys(Keys.ARROW_DOWN)
                grade_entry.send_keys(Keys.RETURN)
                # basic_salary
                basic_salary = row["basic_salary"]
                basic_salary_entry = self.driver.find_element(By.NAME, 'BasicSalary')
                basic_salary_entry.click()
                basic_salary_entry.send_keys(Keys.BACKSPACE)
                basic_salary_entry.send_keys(basic_salary)
                # social_security
                drop_downs[18].click()
                for social_security_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if social_security_.text == "Social Security":
                        social_security_.click()
                        break
                # attendance_type
                drop_downs[21].click()
                for attendance_type_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if attendance_type_.text == "Regular":
                        attendance_type_.click()
                        break
                time.sleep(1)
                # shift_name
                self.driver.find_elements(By.CLASS_NAME, "select2-selection__rendered")[22].click()
                self.driver.find_element(By.CLASS_NAME, "select2-search__field").send_keys("Morning 8 H - Roster")
                for shift_name_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if shift_name_.text == "Morning 8 H - Roster":
                        shift_name_.click()
                        break
                # card_id
                self.driver.find_element(By.NAME, 'CardID').send_keys(code)
                # phone_number
                phone_number = row["phone_number"]
                self.driver.find_element(By.NAME, 'lstRequiredFields[0].FieldValueText').send_keys(phone_number)
                # father_name
                self.driver.find_element(By.NAME, 'lstAdditionalField[0].Description').send_keys(father_name)
                # place_of_issue
                place_of_issue = row["place_of_issue"]
                self.driver.find_element(By.NAME, 'lstAdditionalField[1].Description').send_keys(place_of_issue)
                # birth_certificate_no
                birth_certificate_no = row["birth_certificate_no"]
                self.driver.find_element(By.NAME, 'lstAdditionalField[2].Description').send_keys(birth_certificate_no)
                # birth_certificate_serial
                birth_certificate_serial = row["birth_certificate_serial"]
                self.driver.find_element(By.NAME, 'lstAdditionalField[3].Description').send_keys(birth_certificate_serial)
                # military_service
                military_service_dict = {"مرد": "Done/ Exempt", "زن": "Not applicable"}
                drop_downs[22].click()
                for military_service_ in self.driver.find_elements(By.CLASS_NAME, "select2-results__option"):
                    if military_service_.text == military_service_dict[gender]:
                        military_service_.click()
                        break
                # place_of_issue_fa
                place_of_issue_fa = row["place_of_issue_fa"]
                self.driver.find_element(By.NAME, 'lstAdditionalField[5].Description').send_keys(place_of_issue_fa)
                # address_fa
                address_fa = row["address_fa"]
                self.driver.find_element(By.NAME, 'lstAdditionalField[6].Description').send_keys(address_fa)
                # position_fa
                position_fa = row["position_fa"]
                self.driver.find_element(By.NAME, 'lstAdditionalField[7].Description').send_keys(position_fa)
                # Multi Selectors
                multi_selectors = self.driver.find_elements(By.CLASS_NAME, "multiselect")
                # hrbp
                hrbp = row["hrbp"]
                hrbps = hrbp.split(", ")
                multi_selectors[-1].click()
                hrbp_labels = self.driver.find_elements(By.CLASS_NAME, "checkbox")[-4:]
                for hrbp_label in hrbp_labels:
                    if hrbp_label.text in hrbps:
                        hrbp_label.click()
                multi_selectors[-1].send_keys(Keys.ESCAPE)
                # Wait before Submitting the Form
                time.sleep(1)
                # Add Employee Button
                self.driver.find_element(By.ID, "btnsubmit").click()
                # Clicking Yes Button
                self.wait.until(
                    EC.visibility_of_element_located((By.ID, "BT-alert-btnYes"))
                ).click()
                # Wait between Each Iteration
                time.sleep(3)
            # Stopping the Process
            self.stop_process()
        except:
            self.stop_process()
            messagebox.showerror("Network Error", "Network is unstable.")

    def stop_process(self):
        self.driver.quit()
        try:
            self.pin_window.destroy()
        except AttributeError:
            pass
        self.start_button.config(text="Start Process", command=self.start_process)

if __name__ == "__main__":
    root = tk.Tk()
    app = DataImportApp(root)
    root.mainloop()