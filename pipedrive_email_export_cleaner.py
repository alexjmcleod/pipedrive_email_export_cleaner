# Usage notes
# To run script: 
# python pipedrive_email_export_cleaner.py [input_file_name] [output_file_name]

# Other requirements:
# - Email field must be called "Email" in the input csv
# - The output file must not already exist. This is to prevent any accidental overwrites. 

import csv
import os
import sys

def start():
    # Capture input/output file names
    input_file = sys.argv[1] if len(sys.argv) > 1 else "input_file.csv"
    output_file = sys.argv[2] if len(sys.argv) > 2 else "output_file.csv"
    
    error_check(input_file, output_file)

    process_rows(input_file, output_file)

def process_rows(input_file, output_file):

    # Create dict to track unique contacts
    unique_emails = {}

    with open(input_file, 'r', encoding='utf-8-sig') as readfile:
        csvreader = csv.DictReader(readfile)
        fieldnames = csvreader.fieldnames

        with open(output_file, 'w', newline="", encoding='utf-8-sig') as writefile:
            csvwriter = csv.DictWriter(writefile, fieldnames=fieldnames)
            
            # Write header row
            csvwriter.writeheader()

            for row in csvreader: 
                # Capture the email field
                email_field = row['Email']

                # Skip this row if the email field is blank
                if email_field == "":
                    continue

                # Separate any comma-separated emails and strip to clean
                # Note: This is a list comprehension and just runs each
                # list entry through strip() after splitting. A simpler (to understand)
                # way to accomplish the same thing is like this:
                # separated_emails = email_field.split(',')
                # And then you'd have to strip with an additional loop here 
                # or within the loop below
                separated_emails = [email.strip() for email in email_field.split(',')]

                # For each comma-separated email,
                # check if this email has already been processed.
                # If not, duplicate the original row and place this 
                # email in the row
                for email in separated_emails:
                    if email not in unique_emails:
                        unique_emails[email] = True # Log that we've seen this email
                        
                        new_row = row.copy()
                        new_row['Email'] = email 

                        # Strip each field in the row to clean it
                        for field in new_row.keys():
                            new_row[field] = new_row[field].strip()

                        print("Writing...")
                        print(new_row)
                        print()

                        csvwriter.writerow(new_row)

def error_check(input_file, output_file):
    # Input file is a csv?
    file_name, ext = os.path.splitext(input_file)
    if ext != ".csv":
        sys.exit("Input file must be a CSV.")

    # Output file is a csv?
    file_name, ext = os.path.splitext(output_file)
    if ext != ".csv":
        sys.exit("Output file must be a CSV.")

    # Output file will not be overwriting anything?
    if os.path.exists(output_file):
        sys.exit(f"Output filename {output_file} is in use.\n\nChoose a different output filename or remove the existing file and try again.\n")
    
if __name__ == "__main__":
    start()
