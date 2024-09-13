import pandas as pd

def check_email_in_list(file_path):
    try:
        # loads the spreadsheet XLSX 
        df = pd.read_excel(file_path)
        
        # check if 'email' column exists
        if 'Email' not in df.columns:
            return "The spreadsheet doesn't have an 'Email' column."

        while True:
            # prompt the user for an email
            input_email = input("Enter email: ")

            # check if the user wants to exit
            if input_email.lower() == 'exit':
                print("Exiting the program.")
                break

            # check if the email is in the list
            if input_email in df['Email'].values:
                print(f"Yes, {input_email} is in the list.")
            else:
                print(f"NOOOOOO {input_email} ")
    
    except FileNotFoundError:
        print(f"File not found: {file_path}")

file_path = 'tutors_emails.xlsx' 

check_email_in_list(file_path)
