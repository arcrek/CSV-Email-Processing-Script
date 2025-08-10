import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re
import sys
from pathlib import Path


def safe_print(*args, **kwargs):
    """Print function that won't crash when no console is available (exe mode)"""
    try:
        if hasattr(sys, 'stdout') and sys.stdout:
            print(*args, **kwargs)
    except:
        # Silently ignore print errors when running as exe without console
        pass


def select_csv_file():
    """Open file dialog to select CSV file"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    file_path = filedialog.askopenfilename(
        title="Select CSV file to process",
        filetypes=[
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ]
    )
    
    root.destroy()
    return file_path


def select_txt_file():
    """Open optional file dialog for exclusion list"""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    result = messagebox.askyesno(
        "Optional Exclusion List",
        "Do you want to select a TXT file with emails to exclude?"
    )
    
    if result:
        file_path = filedialog.askopenfilename(
            title="Select TXT file with emails to exclude (Optional)",
            filetypes=[
                ("Text files", "*.txt"),
                ("All files", "*.*")
            ]
        )
        root.destroy()
        return file_path
    else:
        root.destroy()
        return None


def validate_csv_format(df):
    """Validate that CSV contains required columns"""
    required_columns = [
        "Email Address [Required]",
        "Last Sign In [READ ONLY]"
    ]
    
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        raise ValueError(f"Missing required columns: {missing_columns}")
    
    return True


def validate_email(email):
    """Basic email validation"""
    if not isinstance(email, str):
        return False
    
    # Basic email regex pattern
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None


def extract_domain(email):
    """Extract domain from email address"""
    try:
        if '@' in email:
            domain = email.split('@')[1]
            # Clean domain name for filename (remove special characters)
            domain = re.sub(r'[^\w.-]', '', domain)
            return domain
        else:
            return 'unknown'
    except:
        return 'unknown'


def load_exclusion_emails(txt_path):
    """Load emails from TXT file for exclusion"""
    exclusion_emails = set()
    
    try:
        with open(txt_path, 'r', encoding='utf-8') as file:
            for line in file:
                line = line.strip()
                if line:
                    # Try to extract email from line (in case line contains other text)
                    email_match = re.search(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}', line)
                    if email_match:
                        exclusion_emails.add(email_match.group().lower())
                    elif validate_email(line):
                        exclusion_emails.add(line.lower())
        
        safe_print(f"Loaded {len(exclusion_emails)} emails for exclusion")
        return exclusion_emails
        
    except Exception as e:
        safe_print(f"Error reading exclusion file: {e}")
        return set()


def process_csv(csv_path, txt_path=None):
    """Main processing logic"""
    try:
        # Load CSV file
        safe_print(f"Loading CSV file: {csv_path}")
        df = pd.read_csv(csv_path)
        safe_print(f"Loaded {len(df)} rows from CSV")
        
        # Validate CSV format
        validate_csv_format(df)
        
        # Phase 1: Filter by Login Status
        # Keep users who have logged in (NOT "Never logged in") - these are active accounts to delete
        initial_count = len(df)
        df_filtered = df[df['Last Sign In [READ ONLY]'] != 'Never logged in'].copy()
        filtered_count = len(df_filtered)
        
        safe_print(f"Phase 1: Filtered {initial_count - filtered_count} rows (kept {filtered_count} users who have logged in)")
        
        if filtered_count == 0:
            messagebox.showwarning(
                "No Data",
                "No users who have logged in found. Output will be empty."
            )
            return None, None, None
        
        # Phase 2: Optional Email Exclusion
        if txt_path:
            exclusion_emails = load_exclusion_emails(txt_path)
            if exclusion_emails:
                # Convert email column to lowercase for comparison
                email_column = 'Email Address [Required]'
                before_exclusion = len(df_filtered)
                
                # Filter out emails that are in the exclusion list
                df_filtered = df_filtered[
                    ~df_filtered[email_column].str.lower().isin(exclusion_emails)
                ].copy()
                
                after_exclusion = len(df_filtered)
                excluded_count = before_exclusion - after_exclusion
                safe_print(f"Phase 2: Excluded {excluded_count} emails from TXT file")
        
        # Extract emails and validate
        email_column = 'Email Address [Required]'
        emails = df_filtered[email_column].tolist()
        
        # Validate emails and filter out invalid ones
        valid_emails = []
        invalid_count = 0
        
        for email in emails:
            if validate_email(email):
                valid_emails.append(email)
            else:
                invalid_count += 1
                safe_print(f"Warning: Invalid email format skipped: {email}")
        
        if invalid_count > 0:
            safe_print(f"Skipped {invalid_count} invalid email addresses")
        
        if not valid_emails:
            messagebox.showwarning(
                "No Valid Emails",
                "No valid email addresses found after processing."
            )
            return None, None, None
        
        # Extract domain (assume all emails use same domain)
        domain = extract_domain(valid_emails[0])
        safe_print(f"Detected domain: {domain}")
        
        # Verify all emails use the same domain
        domains = set(extract_domain(email) for email in valid_emails)
        if len(domains) > 1:
            safe_print(f"Warning: Multiple domains detected: {domains}")
            safe_print(f"Using primary domain: {domain}")
        
        return valid_emails, domain, csv_path
        
    except Exception as e:
        messagebox.showerror("Error", f"Error processing CSV: {str(e)}")
        return None, None, None


def save_output(emails, domain, csv_path):
    """Save processed emails to output CSV"""
    try:
        # Create output filename
        output_filename = f"{domain}-to-delete.csv"
        
        # Save in same directory as input CSV
        input_dir = Path(csv_path).parent
        output_path = input_dir / output_filename
        
        # Create output DataFrame
        output_df = pd.DataFrame({
            'primaryEmail': emails
        })
        
        # Save to CSV
        output_df.to_csv(output_path, index=False)
        
        safe_print(f"Output saved to: {output_path}")
        safe_print(f"Total emails in output: {len(emails)}")
        
        messagebox.showinfo(
            "Success",
            f"Processing complete!\n\n"
            f"Output file: {output_filename}\n"
            f"Location: {input_dir}\n"
            f"Total emails: {len(emails)}"
        )
        
        return output_path
        
    except Exception as e:
        messagebox.showerror("Error", f"Error saving output: {str(e)}")
        return None


def main():
    """Main execution flow"""
    safe_print("CSV Email Processing Script")
    safe_print("=" * 40)
    
    try:
        # Step 1: Select CSV file
        csv_path = select_csv_file()
        if not csv_path:
            safe_print("No CSV file selected. Exiting.")
            return
        
        # Step 2: Select optional TXT file
        txt_path = select_txt_file()
        if txt_path:
            safe_print(f"TXT file selected: {txt_path}")
        else:
            safe_print("No TXT file selected (skipping exclusions)")
        
        # Step 3: Process CSV
        emails, domain, input_path = process_csv(csv_path, txt_path)
        
        if emails is None:
            safe_print("Processing failed or no data to output.")
            return
        
        # Step 4: Save output
        output_path = save_output(emails, domain, input_path)
        
        if output_path:
            safe_print("Script completed successfully!")
        else:
            safe_print("Script completed with errors.")
            
    except Exception as e:
        safe_print(f"Unexpected error: {e}")
        messagebox.showerror("Unexpected Error", f"An unexpected error occurred: {str(e)}")


if __name__ == "__main__":
    main()
