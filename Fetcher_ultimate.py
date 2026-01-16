import tkinter as tk
from tkinter import scrolledtext, ttk, messagebox, filedialog
import sys
import io
import queue
import threading
import requests
import time
import pandas as pd
from bs4 import BeautifulSoup
import os
import re
import csv
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import subprocess
import importlib

# SMTP Configuration
SMTP_SERVER = 'smtp.gmail.com'  # Change if using another provider
SMTP_PORT = 587
SMTP_USER = 'sender@gmail.com'  # <-- CHANGE THIS
SMTP_PASSWORD = 'password'  # <-- CHANGE THIS (use app password for Gmail)
EMAIL_DOMAIN = 'stu.ucsc.cmb.ac.lk'  # Email domain for students

# Check for required packages
def check_dependencies():
    missing_packages = []
    try:
        import openpyxl
    except ImportError:
        missing_packages.append('openpyxl')
    
    try:
        import requests
    except ImportError:
        missing_packages.append('requests')
    
    try:
        import bs4
    except ImportError:
        missing_packages.append('beautifulsoup4')
    
    if missing_packages:
        print(f"Missing required packages: {', '.join(missing_packages)}")
        print("Installing missing packages...")
        
        try:
            import subprocess
            for package in missing_packages:
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print("All missing packages installed successfully!")
        except Exception as e:
            print(f"Error installing packages: {e}")
            print("Please install the missing packages manually using:")
            print(f"pip install {' '.join(missing_packages)}")
            return False
    return True

# Default credentials
DEFAULT_INDEX = '23000000'
DEFAULT_NIC = '200300000000'
LOGIN_URL = 'https://ucsc.cmb.ac.lk/student-portal/public/index.php/login'
DASHBOARD_URL = 'https://ucsc.cmb.ac.lk/student-portal/public/index.php'
RESULTS_URL = 'https://ucsc.cmb.ac.lk/student-portal/public/index.php/results'
EXCEL_PATH = 'results.xlsx'
CSV_PATH = 'student_credentials.csv'  # Expected format: index,nic,email
CREDIT_CSV_PATH = 'credits.csv'  # Expected format: subject_code,credits

# Email sending function
def send_email(to_email, subject, body):
    """Send an email with the given subject and body to the specified recipient"""
    try:
        print(f"Sending email to: {to_email}")
        print(f"Subject: {subject}")
        
        msg = MIMEMultipart()
        msg['From'] = SMTP_USER
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))
        
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.send_message(msg)
        
        print(f"✓ Email sent successfully to {to_email}")
        return True
    except Exception as e:
        print(f"❌ Error sending email: {e}")
        return False

# GPA Grade Points lookup table
GRADE_POINTS = {
    'A+': 4.00,
    'A': 4.00,
    'A-': 3.70,
    'B+': 3.30,
    'B': 3.00,
    'B-': 2.70,
    'C+': 2.30,
    'C': 2.00,
    'C-': 1.70,
    'D+': 1.30,
    'D': 1.00,
    'E': 0.00,
    'F': 0.00,
    # 'WH':0.00, # WH (Withheld) handled as F in calculate_gpa function (0.00 points, credits counted)
    # 'MC': 0.00  # MC (Medical) handled as F in calculate_gpa function (0.00 points, credits counted)
    # Special cases:
    # MC handled as F in calculate_gpa function (0.00 points, credits counted)
    # NC, CM not considered for GPA calculation
}

class RedirectText(io.StringIO):
    """Class to redirect stdout to the tkinter text widget"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def write(self, string):
        self.text_widget.configure(state=tk.NORMAL)
        self.text_widget.insert(tk.END, string)
        self.text_widget.see(tk.END)  # Auto-scroll to the bottom
        self.text_widget.configure(state=tk.DISABLED)
        
    def flush(self):
        pass

class SimpleLogger:
    # Log level: 0=none, 1=error only, 2=error+warning, 3=all
    log_level = 2
    
    @staticmethod
    def info(message):
        if SimpleLogger.log_level >= 3:
            print(f"[INFO] {message}")
        
    @staticmethod
    def success(message):
        print(f"✓ {message}")
        
    @staticmethod
    def warning(message):
        if SimpleLogger.log_level >= 2:
            print(f"⚠️ {message}")
        
    @staticmethod
    def error(message):
        if SimpleLogger.log_level >= 1:
            print(f"❌ {message}")

class ResultFetcherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("UCSC Result Fetcher (Curl Edition)")
        self.root.geometry("900x700")  # Increased size to accommodate more UI elements
        self.root.resizable(True, True)
        
        # Flag to track running status
        self.running = False
        self.stop_requested = False
          # Email toggle flag
        self.send_emails_enabled = tk.BooleanVar(value=False)
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook/tabs for organization
        self.tabs = ttk.Notebook(main_frame)
        self.tabs.pack(fill=tk.BOTH, expand=True)
        
        # Create tab frames
        main_tab = ttk.Frame(self.tabs, padding="10")
        settings_tab = ttk.Frame(self.tabs, padding="10")
        
        # Add tabs to notebook
        self.tabs.add(main_tab, text="Results Fetcher")
        self.tabs.add(settings_tab, text="SMTP Settings")
        
        # Create login frame
        login_frame = ttk.LabelFrame(main_tab, text="Student Information", padding="10")
        login_frame.pack(fill=tk.X, pady=10)
        
        # Status frame for showing current operation
        status_frame = ttk.Frame(main_tab)
        status_frame.pack(fill=tk.X, pady=5)
        
        # Status label
        ttk.Label(status_frame, text="Status:").pack(side=tk.LEFT, padx=5)
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(status_frame, textvariable=self.status_var, foreground="blue", font=("Arial", 10, "bold"))
        self.status_label.pack(side=tk.LEFT, padx=5)
        
        # Top row for single student entry
        ttk.Label(login_frame, text="Index Number:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.index_var = tk.StringVar(value=DEFAULT_INDEX)
        ttk.Entry(login_frame, textvariable=self.index_var, width=30).grid(row=0, column=1, pady=5, padx=5)
        
        ttk.Label(login_frame, text="NIC Number:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        self.nic_var = tk.StringVar(value=DEFAULT_NIC)
        ttk.Entry(login_frame, textvariable=self.nic_var, width=30).grid(row=1, column=1, pady=5, padx=5)
        
        # Bottom row for file-related inputs
        ttk.Label(login_frame, text="Excel Output:").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5)
        self.excel_var = tk.StringVar(value=EXCEL_PATH)
        ttk.Entry(login_frame, textvariable=self.excel_var, width=30).grid(row=2, column=1, pady=5, padx=5)
        
        # CSV file selection row
        ttk.Label(login_frame, text="Student CSV:").grid(row=3, column=0, sticky=tk.W, pady=5, padx=5)
        self.csv_var = tk.StringVar(value=CSV_PATH)
        csv_entry = ttk.Entry(login_frame, textvariable=self.csv_var, width=30)
        csv_entry.grid(row=3, column=1, pady=5, padx=5)
        
        # Browse button for CSV
        browse_button = ttk.Button(login_frame, text="Browse...", command=self.browse_csv)
        browse_button.grid(row=3, column=2, pady=5, padx=5)
        
        # Credit CSV file selection row (NEW)
        ttk.Label(login_frame, text="Credit CSV:").grid(row=4, column=0, sticky=tk.W, pady=5, padx=5)
        self.credit_csv_var = tk.StringVar(value=CREDIT_CSV_PATH)
        credit_csv_entry = ttk.Entry(login_frame, textvariable=self.credit_csv_var, width=30)
        credit_csv_entry.grid(row=4, column=1, pady=5, padx=5)
        
        # Browse button for Credit CSV
        credit_browse_button = ttk.Button(login_frame, text="Browse...", command=self.browse_credit_csv)
        credit_browse_button.grid(row=4, column=2, pady=5, padx=5)
        
        # Help text for credit CSV
        help_text = "Credit CSV format: subject_code,credits (e.g., SCS1301,3)"
        help_label = ttk.Label(login_frame, text=help_text, foreground="gray", font=("Arial", 8))
        help_label.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=2, padx=5)
        
        # Create a frame for buttons
        button_frame = ttk.Frame(login_frame)
        button_frame.grid(row=6, column=0, columnspan=3, pady=10)
        
        # Fetch button
        self.fetch_button = ttk.Button(button_frame, text="Fetch Single Result", command=self.start_fetching)
        self.fetch_button.pack(side=tk.LEFT, padx=5)
        
        # Batch process button
        self.batch_button = ttk.Button(button_frame, text="Process CSV Batch", command=self.start_batch_processing)
        self.batch_button.pack(side=tk.LEFT, padx=5)
        
        # Stop button
        self.stop_button = ttk.Button(button_frame, text="Stop", command=self.stop_fetching, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=5)
          # Progress bar frame
        progress_frame = ttk.Frame(main_tab)
        progress_frame.pack(fill=tk.X, pady=5)
        
        # Progress bar
        ttk.Label(progress_frame, text="Progress:").pack(side=tk.LEFT, padx=5)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # Progress label
        self.progress_label = ttk.Label(progress_frame, text="0%")
        self.progress_label.pack(side=tk.LEFT, padx=5)
        
        # Logger frame
        logger_frame = ttk.LabelFrame(main_tab, text="Output Log", padding="10")
        logger_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Create scrolled text widget for logging
        self.log_widget = scrolledtext.ScrolledText(logger_frame, wrap=tk.WORD, height=20)
        self.log_widget.pack(fill=tk.BOTH, expand=True)
        self.log_widget.configure(state=tk.DISABLED)
        
        # --- SMTP SETTINGS TAB ---
        smtp_frame = ttk.LabelFrame(settings_tab, text="Email Configuration", padding="10")
        smtp_frame.pack(fill=tk.X, pady=10, padx=10)
        
        # SMTP server settings
        ttk.Label(smtp_frame, text="SMTP Server:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.smtp_server_var = tk.StringVar(value=SMTP_SERVER)
        ttk.Entry(smtp_frame, textvariable=self.smtp_server_var, width=30).grid(row=0, column=1, pady=5, padx=5)
        
        ttk.Label(smtp_frame, text="SMTP Port:").grid(row=1, column=0, sticky=tk.W, pady=5, padx=5)
        self.smtp_port_var = tk.StringVar(value=str(SMTP_PORT))
        ttk.Entry(smtp_frame, textvariable=self.smtp_port_var, width=30).grid(row=1, column=1, pady=5, padx=5)
        
        ttk.Label(smtp_frame, text="Email Address:").grid(row=2, column=0, sticky=tk.W, pady=5, padx=5)
        self.smtp_user_var = tk.StringVar(value=SMTP_USER)
        ttk.Entry(smtp_frame, textvariable=self.smtp_user_var, width=30).grid(row=2, column=1, pady=5, padx=5)
        
        ttk.Label(smtp_frame, text="Password:").grid(row=3, column=0, sticky=tk.W, pady=5, padx=5)
        self.smtp_pass_var = tk.StringVar(value=SMTP_PASSWORD)
        ttk.Entry(smtp_frame, textvariable=self.smtp_pass_var, width=30, show="*").grid(row=3, column=1, pady=5, padx=5)
        
        ttk.Label(smtp_frame, text="Student Email Domain:").grid(row=4, column=0, sticky=tk.W, pady=5, padx=5)
        self.email_domain_var = tk.StringVar(value=EMAIL_DOMAIN)
        ttk.Entry(smtp_frame, textvariable=self.email_domain_var, width=30).grid(row=4, column=1, pady=5, padx=5)
        
        # Add tooltip/help label
        help_text = "Note: For Gmail, use an 'App Password' instead of your regular password.\nGo to Google Account > Security > App Passwords"
        help_label = ttk.Label(smtp_frame, text=help_text, foreground="gray")
        help_label.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=10, padx=5)
        
        # Email toggle checkbox        self.send_emails_enabled = tk.BooleanVar(value=False)
        email_toggle_frame = ttk.Frame(smtp_frame)
        email_toggle_frame.grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=5, padx=5)
        self.email_toggle = ttk.Checkbutton(
            email_toggle_frame, 
            text="Enable Email Notifications", 
            variable=self.send_emails_enabled,
            onvalue=True,
            offvalue=False
        )
        self.email_toggle.pack(side=tk.LEFT)
        ttk.Label(email_toggle_frame, text="(Uncheck to disable sending emails)", foreground="gray").pack(side=tk.LEFT, padx=5)
        
        # Email test section
        test_frame = ttk.LabelFrame(settings_tab, text="Test Email Configuration", padding="10")
        test_frame.pack(fill=tk.X, pady=10, padx=10)
        
        ttk.Label(test_frame, text="Test Email Address:").grid(row=0, column=0, sticky=tk.W, pady=5, padx=5)
        self.test_email_var = tk.StringVar()
        ttk.Entry(test_frame, textvariable=self.test_email_var, width=30).grid(row=0, column=1, pady=5, padx=5)
        
        # Test and Save buttons
        btn_frame = ttk.Frame(test_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, pady=10)
        
        self.test_btn = ttk.Button(btn_frame, text="Send Test Email", command=self.send_test_email)
        self.test_btn.pack(side=tk.LEFT, padx=5)
        
        self.save_btn = ttk.Button(btn_frame, text="Save Settings", command=self.save_smtp_settings)
        self.save_btn.pack(side=tk.LEFT, padx=5)
        
        # Test result label
        self.test_result_var = tk.StringVar()
        self.test_result_label = ttk.Label(test_frame, textvariable=self.test_result_var, font=("Arial", 10))
        self.test_result_label.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5, padx=5)
        
        # Redirect stdout to the text widget
        self.stdout_redirect = RedirectText(self.log_widget)
        sys.stdout = self.stdout_redirect
        
        # Session for making requests
        self.session = None
        
        # Student results and ranking data
        self.student_results = {}
        
        # Check dependencies when starting
        if not check_dependencies():
            messagebox.showwarning("Missing Dependencies", "Please install the missing dependencies before running.")
    
    def browse_csv(self):
        """Open file dialog to select a CSV file"""
        filename = filedialog.askopenfilename(
            title="Select CSV file with student credentials",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.csv_var.set(filename)
            print(f"Selected CSV file: {filename}")
    
    def browse_credit_csv(self):
        """Open file dialog to select a credit CSV file"""
        filename = filedialog.askopenfilename(
            title="Select CSV file with subject credits",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.credit_csv_var.set(filename)
            print(f"Selected Credit CSV file: {filename}")
    
    def start_fetching(self):
        """Start the fetching process for a single student in a separate thread"""
        index_number = self.index_var.get().strip()
        nic = self.nic_var.get().strip()
        excel_path = self.excel_var.get().strip()
        
        if not index_number or not nic:
            print("❌ Please enter both Index Number and NIC.")
            self.update_status("Missing information: Enter Index and NIC", "red")
            return
        
        # Check dependencies again before starting
        if not check_dependencies():
            print("❌ Missing required dependencies. Please install them first.")
            self.update_status("Missing dependencies", "red")
            return
        
        # Set running state and reset stop flag
        self.running = True
        self.stop_requested = False
        
        # Update button states
        self.fetch_button.configure(state=tk.DISABLED)
        self.batch_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        
        # Clear the log
        self.log_widget.configure(state=tk.NORMAL)
        self.log_widget.delete(1.0, tk.END)
        self.log_widget.configure(state=tk.DISABLED)
        
        # Reset progress bar
        self.progress_var.set(0)
        self.progress_label.config(text="0%")
        
        # Update status
        self.update_status(f"Fetching results for index: {index_number}", "blue")
        print(f"\n=== STARTING FETCH FOR INDEX: {index_number} ===\n")
        
        # Start the process in a separate thread
        threading.Thread(
            target=self.run_fetching_process, 
            args=(index_number, nic, excel_path), 
            daemon=True
        ).start()
    
    def start_batch_processing(self):
        """Start the fetching process for multiple students from CSV file in a separate thread"""
        csv_path = self.csv_var.get().strip()
        excel_path = self.excel_var.get().strip()
        
        if not csv_path or not os.path.exists(csv_path):
            print(f"CSV file not found: {csv_path}")
            return
        
        # Check dependencies again before starting
        if not check_dependencies():
            print("Missing required dependencies. Please install them first.")
            return
        
        # Set running state and reset stop flag
        self.running = True
        self.stop_requested = False
        
        # Update button states
        self.fetch_button.configure(state=tk.DISABLED)
        self.batch_button.configure(state=tk.DISABLED)
        self.stop_button.configure(state=tk.NORMAL)
        
        # Clear the log
        self.log_widget.configure(state=tk.NORMAL)
        self.log_widget.delete(1.0, tk.END)
        self.log_widget.configure(state=tk.DISABLED)
        
        # Reset progress bar
        self.progress_var.set(0)
        self.progress_label.config(text="0%")
        
        # Clear student results
        self.student_results = {}
        
        # Start the process in a separate thread
        threading.Thread(
            target=self.run_batch_processing, 
            args=(csv_path, excel_path), 
            daemon=True
        ).start()
    
    def run_batch_processing(self, csv_path, excel_path):
        """Process multiple students from a CSV file"""
        try:
            # Load credit CSV before processing students
            credit_csv_path = self.credit_csv_var.get().strip()
            credit_mapping = load_credit_csv(credit_csv_path)
            
            if not credit_mapping:
                print("⚠️ No credit mapping loaded. GPA calculations may be inaccurate.")
            
            # Read student credentials from CSV file
            students = []
            try:
                with open(csv_path, 'r') as f:
                    reader = csv.reader(f)
                    header = next(reader, None)  # Skip header row if exists
                    
                    # Check if the header row looks like a header or data
                    if header and len(header) >= 2 and not header[0].isdigit():
                        # It's a header, continue with next row
                        print(f"Found header row: {header}")
                        pass
                    else:
                        # No header or header is actually data, add it to students
                        students.append(header)
                    
                    # Add remaining rows
                    for row in reader:
                        if len(row) >= 2:
                            # Make sure each row has at least 3 elements (index, NIC, email)
                            # If email is missing, use None so we can handle it later
                            if len(row) < 3:
                                print(f"⚠️ Row {row} has no email field. Using index@{EMAIL_DOMAIN} as fallback.")
                                row_with_email = row + [None]  # Add None for missing email
                                students.append(row_with_email)
                            else:
                                students.append(row)
            except Exception as e:
                print(f"Error reading CSV file: {e}")
                return
            
            if not students:
                print("No student records found in the CSV file.")
                return
            
            print(f"Found {len(students)} student records in the CSV file.")
            
            # Initialize progress tracking
            total_students = len(students)
            processed_count = 0
            
            # Process each student
            for i, student in enumerate(students):
                if self.stop_requested:
                    print("Batch processing stopped by user.")
                    break
                
                index_number = student[0].strip()
                nic = student[1].strip()
                
                print(f"\n--- Processing student {i+1}/{total_students}: Index {index_number} ---")
                
                # Fetch results for this student
                results = fetch_results(index_number, nic, LOGIN_URL, self)
                if results:
                    # Store student and result
                    student_name = extract_student_name(results) or f"Student {index_number}"
                    
                    # Calculate GPA with credit mapping
                    credits_and_grades = extract_credits_and_grades(results, credit_mapping)
                    gpa, total_credits = calculate_gpa(credits_and_grades)
                          # Get student email from the row (index 2) if available, otherwise generate from index
                    student_email = student[2] if len(student) > 2 and student[2] else f"{index_number}@{EMAIL_DOMAIN}"
                    
                    # Store for ranking, including full results for Excel export
                    self.student_results[student_name] = {
                        'index': index_number,
                        'email': student_email,
                        'gpa': gpa,
                        'total_credits': total_credits,
                        'results': results.copy()  # Store complete results for Excel
                    }
                    
                    print(f"✓ {student_name} (Index: {index_number}) - GPA: {gpa:.2f} (Total Credits: {total_credits})")
                    
                    # Update Excel with results
                    update_excel(index_number, results, excel_path)
                else:
                    print(f"❌ Failed to get results for student with index {index_number}")
                
                # Update progress
                processed_count += 1
                progress = (processed_count / total_students) * 100
                self.root.after(0, self.update_progress, progress)
                
                # Add a small delay to avoid overwhelming the server
                time.sleep(1)
            
            # Display final rankings
            self.display_rankings()
            
        except Exception as e:
            print(f"Unhandled exception in batch processing: {e}")
        finally:
            # Reset running state
            self.running = False
            self.stop_requested = False
            
            # Re-enable buttons
            self.root.after(0, lambda: self.fetch_button.configure(state=tk.NORMAL))
            self.root.after(0, lambda: self.batch_button.configure(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_button.configure(state=tk.DISABLED))
    
    def display_rankings(self):
        """Display the final rankings of students based on GPA, save Excel in rank order, and send emails"""
        if not self.student_results:
            print("\n--- No student results available for ranking ---")
            return
        
        print("\n========== STUDENT RANKINGS BY GPA ==========")
        
        # Sort students by GPA (descending)
        sorted_students = sorted(
            self.student_results.items(), 
            key=lambda x: x[1]['gpa'], 
            reverse=True
        )
        
        # Display ranking and prepare rank data
        ranked_data = []
        for rank, (name, data) in enumerate(sorted_students, 1):
            print(f"Rank {rank}: {name} (Index: {data['index']}) - GPA: {data['gpa']:.2f} (Credits: {data['total_credits']})")
            data['rank'] = rank
            ranked_data.append((rank, name, data))
        
        print("=============================================")
        
        # Save Excel in rank order
        self.save_ranked_excel(ranked_data)
        
        # Send emails to students
        self.send_rank_emails(ranked_data)
    
    def save_ranked_excel(self, ranked_data):
        """Save results to Excel in rank order with GPA and all results"""
        excel_path = self.excel_var.get().strip()
        print(f"\n--- Saving ranked results to Excel: {excel_path} ---")
        
        # Collect all subjects across all students
        all_subjects = set()
        for _, _, data in ranked_data:
            if 'results' in data:
                # Filter out 'student_name' which isn't a subject
                subjects = {k for k in data['results'].keys() if k != 'student_name'}
                all_subjects.update(subjects)
          # Build DataFrame with rank order
        rows = []
        for rank, name, data in ranked_data:
            row = {
                'Rank': rank,
                'Name': name,
                'Index': data['index'],
                'Email': data.get('email', f"{data['index']}@{EMAIL_DOMAIN}"),
                'GPA': data['gpa'],
                'Credits': data['total_credits']
            }
            
            # Add subject grades if available
            if 'results' in data:
                for subject in all_subjects:
                    row[subject] = data['results'].get(subject, '')
            
            rows.append(row)
          # Create DataFrame and save
        try:
            df = pd.DataFrame(rows)
            df.to_excel(excel_path, index=False)
            print(f"✓ Ranked Excel file saved successfully: {excel_path}")
        except Exception as e:
            print(f"❌ Error saving ranked Excel file: {e}")
            
            # Try saving as CSV as a fallback
            try:
                csv_path = excel_path.replace('.xlsx', '.csv')
                df.to_csv(csv_path)
                print(f"Results saved as CSV instead: {csv_path}")
            except Exception as csv_err:
                print(f"Could not save as CSV either: {csv_err}")
    
    def send_rank_emails(self, ranked_data):
        """Send emails to students with their rank information"""
        # Check if email sending is enabled
        if not self.send_emails_enabled.get():
            print("\n--- Email notifications are disabled ---")
            self.update_status("Email notifications are disabled", "blue")
            return
            
        print("\n--- Sending rank notification emails to students ---")
        self.update_status("Sending rank emails...", "blue")
        
        email_count = 0
        total_emails = len(ranked_data)
        
        for i, (rank, name, data) in enumerate(ranked_data):
            index = data['index']
            # Use the email address from the CSV if available, otherwise fallback to generated email
            to_email = data.get('email')
            if not to_email:
                to_email = f"{index}@{EMAIL_DOMAIN}"
                print(f"⚠️ No email address found for {name} (Index: {index}). Using {to_email} as fallback.")
            
            subject = "Your UCSC Academic Rank Notification"
            
            body = f"""
Dear {name},

This is an automated notification from the UCSC Result Fetcher system.

Based on the latest batch processing of academic results, your current rank is:

Rank: {rank}

This ranking is calculated based on your GPA ({data['gpa']:.2f}) among all processed students.
Please note that this is for informational purposes only and may not reflect the official university standings.

This is an auto-generated email. Please do not reply to this message.

Best regards,
UCSC Result Fetcher System
"""
            # Update progress for email sending
            progress = ((i + 1) / total_emails) * 100
            self.root.after(0, self.update_progress, progress)
            self.update_status(f"Sending email {i+1}/{total_emails}: {name}", "blue")
            
            try:
                if send_email(to_email, subject, body):
                    print(f"✓ Email sent successfully to {name} (Index: {index}, Rank: {rank})")
                    email_count += 1
                else:
                    print(f"❌ Failed to send email to {name} (Index: {index})")
            except Exception as e:
                print(f"❌ Error sending email to {index}: {e}")
            
            # Brief pause to avoid overwhelming the SMTP server
            time.sleep(0.5)
        
        print(f"\n--- Sent {email_count} rank notification emails ---")
        self.update_status(f"Completed: Sent {email_count}/{total_emails} emails", "green" if email_count == total_emails else "orange")
    
    def update_progress(self, progress):
        """Update the progress bar and label"""
        self.progress_var.set(progress)
        self.progress_label.config(text=f"{progress:.1f}%")
    
    def stop_fetching(self):
        """Stop the fetching process"""
        if not self.running:
            return
        
        print("Stopping the process... Please wait.")
        self.stop_requested = True
        self.stop_button.configure(state=tk.DISABLED)
    
    def run_fetching_process(self, index_number, nic, excel_path):
        """Run the main fetching process for a single student"""
        try:
            # Load credit CSV before processing
            credit_csv_path = self.credit_csv_var.get().strip()
            credit_mapping = load_credit_csv(credit_csv_path)
            
            if not credit_mapping:
                print("⚠️ No credit mapping loaded. GPA calculations may be inaccurate.")
            
            # Set progress to indicate start
            self.update_progress(10)
            
            results = fetch_results(index_number, nic, LOGIN_URL, self)
            self.update_progress(50)
            
            if results and not self.stop_requested:
                # Extract student name if available
                student_name = extract_student_name(results) or f"Student {index_number}"
                
                # Calculate GPA with credit mapping
                credits_and_grades = extract_credits_and_grades(results, credit_mapping)
                gpa, total_credits = calculate_gpa(credits_and_grades)
                
                print(f"\n--- Results for {student_name} ---")
                print(f"GPA: {gpa:.2f} (Total Credits: {total_credits})")
                
                # Store for potential ranking, including full results for Excel
                self.student_results[student_name] = {
                    'index': index_number,
                    'email': f"{index_number}@{EMAIL_DOMAIN}",  # Default for single student mode
                    'gpa': gpa,
                    'total_credits': total_credits,
                    'results': results.copy()  # Store complete results for Excel
                }
                
                update_excel(index_number, results, excel_path)
                print(f"Results updated for {index_number}")
                self.update_progress(100)
            else:
                self.update_progress(100)
        except Exception as e:
            print(f"Unhandled exception: {e}")
            self.update_progress(100)
        finally:
            # Reset running state
            self.running = False
            self.stop_requested = False
              # Re-enable the start button and disable stop button when done
            self.root.after(0, lambda: self.fetch_button.configure(state=tk.NORMAL))
            self.root.after(0, lambda: self.batch_button.configure(state=tk.NORMAL))
            self.root.after(0, lambda: self.stop_button.configure(state=tk.DISABLED))
    
    def update_status(self, message, color="blue"):
        """Update the status bar with a message and color"""
        self.status_var.set(message)
        self.status_label.config(foreground=color)
        self.root.update_idletasks()  # Force UI update
    
    def save_smtp_settings(self):
        """Save SMTP settings from the UI to the global variables"""
        global SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, EMAIL_DOMAIN
        
        try:
            # Update global variables with values from UI
            SMTP_SERVER = self.smtp_server_var.get().strip()
            SMTP_PORT = int(self.smtp_port_var.get().strip())
            SMTP_USER = self.smtp_user_var.get().strip()
            SMTP_PASSWORD = self.smtp_pass_var.get().strip()
            EMAIL_DOMAIN = self.email_domain_var.get().strip()
            
            # Update the test result label
            self.test_result_var.set("Settings saved successfully!")
            self.test_result_label.config(foreground="green")
            
            # Also update status
            self.update_status("SMTP settings updated", "green")
            
            print(f"SMTP settings updated: {SMTP_SERVER}:{SMTP_PORT}")
            return True
        except Exception as e:
            self.test_result_var.set(f"Error saving settings: {e}")
            self.test_result_label.config(foreground="red")
            print(f"Error saving SMTP settings: {e}")
            self.update_status(f"Error: {str(e)}", "red")
            return False
    
    def send_test_email(self):
        """Send a test email using the current SMTP settings"""
        test_email = self.test_email_var.get().strip()
        
        if not test_email:
            self.test_result_var.set("Please enter a test email address.")
            self.test_result_label.config(foreground="red")
            return
        
        # Check if email sending is enabled
        if not self.send_emails_enabled.get():
            self.test_result_var.set("Email notifications are disabled. Enable them first.")
            self.test_result_label.config(foreground="orange")
            return
        
        # Save settings first
        if not self.save_smtp_settings():
            return
        
        # Prepare test email
        subject = "UCSC Result Fetcher - Test Email"
        body = f"""
Hello,

This is a test email from the UCSC Result Fetcher application.
If you received this email, your SMTP settings are configured correctly.

SMTP Server: {SMTP_SERVER}
SMTP Port: {SMTP_PORT}
From Email: {SMTP_USER}

This is an automated message. Please do not reply.

Best regards,
UCSC Result Fetcher
"""
        
        # Update UI
        self.test_result_var.set("Sending test email...")
        self.test_result_label.config(foreground="blue")
        self.update_status("Sending test email...", "blue")
        self.root.update()  # Force UI update
        
        # Send email
        try:
            if send_email(to_email=test_email, subject=subject, body=body):
                self.test_result_var.set(f"Test email sent successfully to {test_email}")
                self.test_result_label.config(foreground="green")
                self.update_status("Email sent successfully", "green")
            else:
                self.test_result_var.set("Failed to send test email. Check console for details.")
                self.test_result_label.config(foreground="red")
                self.update_status("Failed to send email", "red")
        except Exception as e:
            self.test_result_var.set(f"Error: {str(e)}")
            self.test_result_label.config(foreground="red")
            print(f"Error sending test email: {e}")
            self.update_status(f"Error: {str(e)}", "red")


def load_credit_csv(csv_path):
    """Load credit mapping from CSV file. Returns dict {subject_code: credits}"""
    credit_mapping = {}
    
    if not csv_path or not os.path.exists(csv_path):
        print(f"⚠️ Credit CSV file not found: {csv_path}")
        return credit_mapping
    
    try:
        with open(csv_path, 'r') as f:
            reader = csv.reader(f)
            header = next(reader, None)  # Skip header if exists
            
            # Check if header looks like actual data
            if header and len(header) >= 2:
                try:
                    # If second column is a number, it's data not header
                    float(header[1])
                    credit_mapping[header[0].strip()] = float(header[1])
                except ValueError:
                    # It's a header, skip it
                    pass
            
            # Process remaining rows
            for row in reader:
                if len(row) >= 2:
                    subject_code = row[0].strip()
                    try:
                        credits = float(row[1].strip())
                        credit_mapping[subject_code] = credits
                    except ValueError:
                        print(f"⚠️ Invalid credit value for {subject_code}: {row[1]}")
        
        print(f"✓ Loaded {len(credit_mapping)} subject credits from {csv_path}")
        # Debug: Show first few keys
        if credit_mapping:
            sample_keys = list(credit_mapping.keys())[:5]
            print(f"DEBUG: Sample credit keys: {sample_keys}")
        return credit_mapping
        
    except Exception as e:
        print(f"❌ Error loading credit CSV: {e}")
        return credit_mapping

def extract_student_name(results):
    """Extract student name from results if available"""
    if isinstance(results, dict) and 'student_name' in results:
        return results['student_name']
    return None

def extract_credits_and_grades(results, credit_mapping=None):
    """Extract credits and grades from results using external credit mapping"""
    # First, process the results to handle duplicate subjects (like MC retakes)
    processed_subjects = {}
    
    if credit_mapping is None:
        credit_mapping = {}
    
    for subject_with_details, grade in results.items():
        # Skip non-subject entries (like student_name)
        if subject_with_details == 'student_name':
            continue
            
        # Extract subject code (e.g., "SCS1301" from "SCS1301 Data Structures")
        # Subject code is typically at the start and is alphanumeric
        subject_code_match = re.match(r'([A-Z]{3}\d{4})', subject_with_details.strip())
        if not subject_code_match:
            print(f"⚠️ Could not extract subject code from: {subject_with_details}")
            continue
        
        subject_code = subject_code_match.group(1)
        
        # Store the grade in processed_subjects, but prefer non-MC grades
        if subject_code in processed_subjects:
            existing_grade = processed_subjects[subject_code]
            # If existing entry is MC and current is not MC, replace it
            if existing_grade.strip().upper() == 'MC' and grade.strip().upper() != 'MC':
                processed_subjects[subject_code] = grade
        else:
            processed_subjects[subject_code] = grade
    
    # Now create credits_and_grades using credit mapping
    credits_and_grades = []
    
    for subject_code, grade in processed_subjects.items():
        # Look up credits from mapping
        if subject_code in credit_mapping:
            credits = credit_mapping[subject_code]
        else:
            # Debug: Show what we're looking for
            print(f"⚠️ No credit mapping found for '{subject_code}' (looking in {len(credit_mapping)} keys)")
            credits = 1.0
        
        # Store (credits, grade, subject_code) for calculate_gpa
        credits_and_grades.append((credits, grade, subject_code))
    
    return credits_and_grades

def calculate_gpa(credits_and_grades):
    """Calculate GPA from list of (credits, grade, subject) tuples"""
    total_grade_points = 0.0
    total_credits = 0.0
    
    for credits, grade, subject in credits_and_grades:
        # Clean up the grade (remove spaces, convert to uppercase)
        clean_grade = grade.strip().upper()
        # Skip ENH subjects (Enhancement subjects)
        if subject.strip().upper().startswith('EN'):
            # print(f"Skipping ENH subject in GPA calculation: {subject}")
            continue
        
        # Handle special cases:
        # NC, CM - Not considered for GPA (neither credits nor grade points)
        # MC - Considered as F (0 grade points, but credits are counted)

        # Skip NC and CM grades
        elif clean_grade == 'NC' or clean_grade == 'CM':
            continue
        elif clean_grade == 'MC' or clean_grade == 'WH':
            # MC is treated as an F (0 points, but credits are counted)
            # WH is treated as a withdrawal (0 points, but credits are counted)
            total_grade_points += 0.0
            total_credits += credits
        # Regular grades from the lookup table
        elif clean_grade in GRADE_POINTS:
            grade_point = GRADE_POINTS[clean_grade]
            total_grade_points += credits * grade_point
            total_credits += credits
    
    # Calculate GPA
    if total_credits > 0:
        gpa = total_grade_points / total_credits
    else:
        gpa = 0.0
    
    return gpa, total_credits

def fetch_results(index_number, nic, login_url, gui_instance=None):
    """Function to log in and fetch results table using requests instead of Selenium"""
    logger = SimpleLogger()
    
    # Create a session to maintain cookies
    logger.info("Initializing HTTP session...")
    session = requests.Session()
    
    # Set headers to mimic a browser
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
    }
    
    try:
        # Get the initial page to extract any CSRF token if needed
        logger.info(f"Connecting to {login_url}...")
        response = session.get(login_url, headers=headers)
        response.raise_for_status()
        
        # Check for stop request
        if gui_instance and gui_instance.stop_requested:
            logger.info("Process terminated by user")
            return None
        
        # Parse the HTML to identify the form and check for hidden fields
        soup = BeautifulSoup(response.text, 'html.parser')
        form = soup.find('form')
        
        if not form:
            logger.error("Could not find form on the page.")
            return None
        
        # Get the form action URL
        action = form.get('action', '')
        if not action:
            action = login_url
        elif not action.startswith('http'):
            # Handle relative URLs
            if action.startswith('/'):
                base_url = '/'.join(login_url.split('/')[:3])  # http(s)://domain
                action = base_url + action
            else:
                action = login_url.rstrip('/') + '/' + action
        
        logger.info(f"Form action URL: {action}")
          # Prepare form data
        form_data = {}
        hidden_inputs = form.find_all('input', type='hidden')
        for hidden in hidden_inputs:
            name = hidden.get('name', '')
            value = hidden.get('value', '')
            if name:
                form_data[name] = value
        
        # Add the user credentials
        input_fields = form.find_all('input', type=['text', 'password'])
        if len(input_fields) >= 2:
            index_field_name = input_fields[0].get('name', 'index')
            nic_field_name = input_fields[1].get('name', 'nic')
            form_data[index_field_name] = index_number
            form_data[nic_field_name] = nic
        else:
            # Default field names
            form_data['index'] = index_number
            form_data['nic'] = nic
        
        logger.info("Submitting credentials...")
        
        # Submit the form
        response = session.post(action, data=form_data, headers=headers, allow_redirects=True)
        response.raise_for_status()
        
        # Check for stop request
        if gui_instance and gui_instance.stop_requested:
            logger.info("Process terminated by user")
            return None
        
        # Step 2: After login, navigate to dashboard to get results link
        logger.info("Login successful, navigating to dashboard...")
        response = session.get(DASHBOARD_URL, headers=headers)
        response.raise_for_status()
        
        # Check for stop request
        if gui_instance and gui_instance.stop_requested:
            logger.info("Process terminated by user")
            return None
        
        # Step 3: Navigate to results page
        logger.info("Navigating to results page...")
        response = session.get(RESULTS_URL, headers=headers)
        response.raise_for_status()
        
        # Check for stop request
        if gui_instance and gui_instance.stop_requested:
            logger.info("Process terminated by user")
            return None
          # Now parse the results page
        logger.info("Looking for results table...")
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Extract student name if available
        student_name = None
        name_element = soup.find('h3')
        if name_element:
            name_text = name_element.get_text(strip=True)
            name_match = re.search(r'Name\s*:\s*([^,]+)', name_text)
            if name_match:
                student_name = name_match.group(1).strip()
                logger.info(f"Found student name: {student_name}")
        
        # Find ALL results tables (one for each year)
        tables = soup.find_all('table', {'class': 'table-bordered'})
        
        results = {}
        if student_name:
            results['student_name'] = student_name
            
        if tables:
            logger.success(f"Found {len(tables)} results table(s)!")
            
            # Process each table (each year)
            for table_index, table in enumerate(tables):
                rows = table.find_all('tr')[1:]  # Skip header row
                for row in rows:
                    cols = row.find_all('td')
                    if len(cols) >= 5:  # Ensure we have at least 5 columns (subject, year, semester, credits, result)
                        subject_with_name = cols[0].text.strip()
                        grade = cols[4].text.strip()
                        
                        # Store only subject code and grade (no credits from HTML)
                        # Credits will be looked up from CSV later
                        results[subject_with_name] = grade
        else:
            # If no table found, check for login errors
            error_msg = soup.find('div', {'class': 'alert-danger'})
            if error_msg:
                logger.error(f"Login error: {error_msg.text.strip()}")
            else:
                logger.error("Results table not found. Please check your credentials.")

        return results

    except requests.exceptions.RequestException as e:
        logger.error(f"Network error: {e}")
        return None
    except Exception as e:
        logger.error(f"Error fetching results: {e}")
        return None

def update_excel(index_number, results, excel_path):
    """Function to update Excel file with fetched results"""
    try:
        # Make sure we have openpyxl installed
        try:
            import openpyxl
        except ImportError:
            print("The openpyxl module is not installed. Attempting to install it now...")
            import subprocess
            subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
            print("openpyxl installed successfully")
            # Reimport pandas to ensure it now uses openpyxl
            import importlib
            importlib.reload(pd)
            
        # Try to read existing Excel file
        try:
            if os.path.exists(excel_path):
                df = pd.read_excel(excel_path, index_col=0)
                print(f"Loaded existing Excel file: {excel_path}")
            else:
                # Create new DataFrame if file doesn't exist
                df = pd.DataFrame()
                print(f"Creating new Excel file: {excel_path}")
        except Exception as file_error:
            print(f"Error reading Excel file: {file_error}")
            # Create new DataFrame
            df = pd.DataFrame()
            print(f"Creating new Excel file: {excel_path}")
        
        # Add new columns for any new subjects
        for subject in results:
            if subject not in df.columns:
                df[subject] = ''
                print(f"Added new column: {subject}")
        
        # Update the row for this student
        df.loc[index_number] = [results.get(col, '') for col in df.columns]
        print(f"Updated results for index: {index_number}")
        
        # Save to Excel - with error handling
        try:
            df.to_excel(excel_path)
            print(f"Excel file saved successfully: {excel_path}")
            return True
        except Exception as save_error:
            print(f"Error saving Excel file: {save_error}")
            
            # Try saving as CSV as a fallback
            try:
                csv_path = excel_path.replace('.xlsx', '.csv')
                df.to_csv(csv_path)
                print(f"Results saved as CSV instead: {csv_path}")
                return True
            except:
                print("Could not save as CSV either.")
                return False
                
    except Exception as e:
        print(f"Error updating Excel file: {e}")
        return False

def run_standalone(index_number, nic, excel_path):
    """Run in standalone mode (without GUI)"""
    # Check dependencies
    check_dependencies()
    
    print(f"Fetching results for index: {index_number}")
    results = fetch_results(index_number, nic, LOGIN_URL)
    if results:
        # Calculate GPA if results found
        student_name = extract_student_name(results) or f"Student {index_number}"
        credits_and_grades = extract_credits_and_grades(results)
        gpa, total_credits = calculate_gpa(credits_and_grades)
        print(f"\n--- Results for {student_name} ---")
        print(f"GPA: {gpa:.2f} (Total Credits: {total_credits})")
        
        update_excel(index_number, results, excel_path)
        print(f"Results updated for {index_number}")
    else:
        print("No results found or error occurred.")

if __name__ == "__main__":
    # Check if any command line arguments were provided
    if len(sys.argv) > 1 and sys.argv[1] == "--nogui":
        # Run in standalone mode
        run_standalone(DEFAULT_INDEX, DEFAULT_NIC, EXCEL_PATH)
    else:
        # Run with GUI
        root = tk.Tk()
        app = ResultFetcherGUI(root)
        
        # Start the main event loop
        root.mainloop()
        
        # Restore stdout when the application closes
        sys.stdout = sys.__stdout__