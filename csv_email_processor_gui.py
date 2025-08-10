#!/usr/bin/env python3
"""
CSV Email Processing Script - GUI Version

Simple GUI interface for processing CSV files to filter and extract email addresses.
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re
from pathlib import Path
import threading


class CSVEmailProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Email Processor")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Variables
        self.csv_path = None
        self.txt_path = None
        
        self.setup_ui()
    
    def setup_ui(self):
        """Setup the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="CSV Email Processor", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=4, pady=(0, 20))
        
        # CSV File Selection
        ttk.Label(main_frame, text="CSV File:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.csv_label = ttk.Label(main_frame, text="No file selected", 
                                  background="white", relief="sunken")
        self.csv_label.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        self.csv_button = ttk.Button(main_frame, text="Select CSV", 
                                    command=self.select_csv_file)
        self.csv_button.grid(row=1, column=2, padx=(0, 5), pady=5)
        
        self.csv_clear_button = ttk.Button(main_frame, text="Clear", 
                                          command=self.clear_csv_file, state="disabled")
        self.csv_clear_button.grid(row=1, column=3, pady=5)
        
        # TXT File Selection (Optional)
        ttk.Label(main_frame, text="Emails to Keep:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.txt_label = ttk.Label(main_frame, text="No file selected (optional)", 
                                  background="white", relief="sunken")
        self.txt_label.grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 5), pady=5)
        
        self.txt_button = ttk.Button(main_frame, text="Select TXT", 
                                    command=self.select_txt_file)
        self.txt_button.grid(row=2, column=2, padx=(0, 5), pady=5)
        
        self.txt_clear_button = ttk.Button(main_frame, text="Clear", 
                                          command=self.clear_txt_file, state="disabled")
        self.txt_clear_button.grid(row=2, column=3, pady=5)
        
        # Start Button
        self.start_button = ttk.Button(main_frame, text="Start Processing", 
                                      command=self.start_processing, 
                                      style="Accent.TButton")
        self.start_button.grid(row=3, column=0, columnspan=4, pady=20)
        
        # Status Log
        ttk.Label(main_frame, text="Status Log:").grid(row=4, column=0, sticky=tk.W)
        
        # Text area with scrollbar
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=5, column=0, columnspan=4, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(5, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(log_frame, height=15, width=70)
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Initial log message
        self.log("Welcome to CSV Email Processor!")
        self.log("1. Select a CSV file (required)")
        self.log("2. Optionally select a TXT file for exclusions")
        self.log("3. Use 'Clear' buttons to remove selections if needed")
        self.log("4. Click 'Start Processing' when ready")
        self.log("5. Watch this log for processing updates")
        self.log("-" * 60)
    
    def log(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def select_csv_file(self):
        """Select CSV file"""
        file_path = filedialog.askopenfilename(
            title="Select CSV file to process",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            self.csv_path = file_path
            filename = os.path.basename(file_path)
            self.csv_label.config(text=filename)
            self.csv_clear_button.config(state="normal")
            self.log(f"Selected CSV: {filename}")
        else:
            self.csv_path = None
            self.csv_label.config(text="No file selected")
            self.csv_clear_button.config(state="disabled")
    
    def select_txt_file(self):
        """Select TXT file (optional)"""
        file_path = filedialog.askopenfilename(
            title="Select TXT file with emails to keep (Optional)",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        
        if file_path:
            self.txt_path = file_path
            filename = os.path.basename(file_path)
            self.txt_label.config(text=filename)
            self.txt_clear_button.config(state="normal")
            self.log(f"Selected TXT: {filename}")
        else:
            self.txt_path = None
            self.txt_label.config(text="No file selected (optional)")
            self.txt_clear_button.config(state="disabled")
    
    def clear_csv_file(self):
        """Clear CSV file selection"""
        self.csv_path = None
        self.csv_label.config(text="No file selected")
        self.csv_clear_button.config(state="disabled")
        self.log("Cleared CSV file selection")
    
    def clear_txt_file(self):
        """Clear TXT file selection"""
        self.txt_path = None
        self.txt_label.config(text="No file selected (optional)")
        self.txt_clear_button.config(state="disabled")
        self.log("Cleared TXT file selection")
    
    def start_processing(self):
        """Start processing in a separate thread"""
        if not self.csv_path:
            messagebox.showerror("Error", "Please select a CSV file first!")
            return
        
        # Start processing in thread to prevent UI freezing
        self.start_button.config(state="disabled")
        
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
    
    def process_files(self):
        """Process the selected files"""
        try:
            self.log("Starting processing...")
            self.log(f"CSV file: {os.path.basename(self.csv_path)}")
            if self.txt_path:
                self.log(f"TXT file: {os.path.basename(self.txt_path)}")
            else:
                self.log("No TXT file selected (no exclusions)")
            self.log("-" * 40)
            
            # Load and validate CSV
            self.log("Loading CSV file...")
            df = pd.read_csv(self.csv_path)
            self.log(f"Loaded {len(df)} rows from CSV")
            
            # Validate required columns
            required_columns = ["Email Address [Required]", "Last Sign In [READ ONLY]"]
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                raise ValueError(f"Missing required columns: {missing_columns}")
            
            # Phase 1: Filter by login status
            initial_count = len(df)
            df_filtered = df[df['Last Sign In [READ ONLY]'] != 'Never logged in'].copy()
            filtered_count = len(df_filtered)
            
            self.log(f"Phase 1: Found {filtered_count} users who have logged in")
            
            if filtered_count == 0:
                self.log("Warning: No users who have logged in found!")
                messagebox.showwarning("No Data", "No users who have logged in found.")
                return
            
            # Phase 2: Optional email exclusion
            if self.txt_path:
                self.log("Loading exclusion emails from TXT file...")
                exclusion_emails = self.load_exclusion_emails(self.txt_path)
                
                if exclusion_emails:
                    before_exclusion = len(df_filtered)
                    email_column = 'Email Address [Required]'
                    
                    # Filter out emails in exclusion list
                    df_filtered = df_filtered[
                        ~df_filtered[email_column].str.lower().isin(exclusion_emails)
                    ].copy()
                    
                    after_exclusion = len(df_filtered)
                    excluded_count = before_exclusion - after_exclusion
                    self.log(f"Phase 2: Excluded {excluded_count} emails from TXT file")
            
            # Extract and validate emails
            emails = df_filtered['Email Address [Required]'].tolist()
            valid_emails = []
            invalid_count = 0
            
            for email in emails:
                if self.validate_email(email):
                    valid_emails.append(email)
                else:
                    invalid_count += 1
                    self.log(f"Warning: Invalid email skipped: {email}")
            
            if invalid_count > 0:
                self.log(f"Skipped {invalid_count} invalid email addresses")
            
            if not valid_emails:
                self.log("Error: No valid emails found!")
                messagebox.showerror("Error", "No valid email addresses found.")
                return
            
            # Extract domain and save output
            domain = self.extract_domain(valid_emails[0])
            self.log(f"Detected domain: {domain}")
            
            # Check for multiple domains
            domains = set(self.extract_domain(email) for email in valid_emails)
            if len(domains) > 1:
                self.log(f"Warning: Multiple domains detected: {domains}")
            
            # Save output file
            output_filename = f"{domain}-to-delete.csv"
            input_dir = Path(self.csv_path).parent
            output_path = input_dir / output_filename
            
            output_df = pd.DataFrame({'primaryEmail': valid_emails})
            output_df.to_csv(output_path, index=False)
            
            self.log(f"✓ Output saved: {output_filename}")
            self.log(f"✓ Location: {input_dir}")
            self.log(f"✓ Total emails: {len(valid_emails)}")
            self.log("-" * 40)
            self.log("Processing completed successfully!")
            
            # Show success message
            self.root.after(0, lambda: messagebox.showinfo(
                "Success", 
                f"Processing complete!\n\n"
                f"Output file: {output_filename}\n"
                f"Location: {input_dir}\n"
                f"Total emails: {len(valid_emails)}"
            ))
            
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            self.log(error_msg)
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))
        
        finally:
            # Re-enable button when finished
            self.root.after(0, self.processing_finished)
    
    def processing_finished(self):
        """Called when processing is finished"""
        self.start_button.config(state="normal")
    
    def load_exclusion_emails(self, txt_path):
        """Load emails from TXT file for exclusion"""
        exclusion_emails = set()
        
        try:
            with open(txt_path, 'r', encoding='utf-8') as file:
                for line in file:
                    line = line.strip()
                    if line:
                        # Try to extract email from line
                        email_match = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', line)
                        if email_match:
                            exclusion_emails.add(email_match.group().lower())
                        elif self.validate_email(line):
                            exclusion_emails.add(line.lower())
            
            self.log(f"Loaded {len(exclusion_emails)} emails for exclusion")
            return exclusion_emails
            
        except Exception as e:
            self.log(f"Error reading exclusion file: {e}")
            return set()
    
    def validate_email(self, email):
        """Basic email validation"""
        if not isinstance(email, str):
            return False
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def extract_domain(self, email):
        """Extract domain from email address"""
        try:
            if '@' in email:
                domain = email.split('@')[1]
                domain = re.sub(r'[^\w.-]', '', domain)
                return domain
            else:
                return 'unknown'
        except:
            return 'unknown'


def main():
    """Main function"""
    root = tk.Tk()
    app = CSVEmailProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
