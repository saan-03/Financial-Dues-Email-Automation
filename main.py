import pandas as pd
import duckdb as db
import win32com.client as win32

# Connecting to virtual database
con = db.connect('connections.db')

# Simulating the financial due data table
finance_due_data = pd.DataFrame({
    'vendor_name': ['Vendor A', 'Vendor B', 'Vendor C'],
    'due_date': ['2025-02-25', '2025-03-05', '2025-03-15'],
    'amount_due': [1000, 1500, 500],
    'invoice_number': ['INV001', 'INV002', 'INV003']
})

# Saving this to the database for simulation
con.register('finance_due_data', finance_due_data)

# Simulating the finance emails DataFrame
finance_emails_df = pd.DataFrame({
    'Vendor': ['Vendor A', 'Vendor B', 'Vendor C'],
    'Emails': ['finance@vendora.com', 'finance@vendorb.com', 'finance@vendorc.com']
})

# Sending Finance Due Emails
def send_finance_due_emails():
    # Query for upcoming financial due data
    finance_due_data = con.sql("""
        SELECT vendor_name, due_date, amount_due, invoice_number
        FROM finance_due_data
        WHERE due_date >= CURRENT_DATE
    """).to_df()

    # Prepare the email content (convert to HTML table)
    finance_due_table = finance_due_data.to_html(index=False, max_rows=5, bold_rows=False)
    
    finance_emails = finance_emails_df['Emails'].iloc[0]
    spec_email_df = con.sql("""
        SELECT DISTINCT finance_contact_email 
        FROM finance_due_data 
        WHERE due_date >= CURRENT_DATE
    """).to_df()
    
    # Remove NaN values from the email list
    spec_email_df.dropna(subset=['finance_contact_email'], inplace=True)
    
    spec_cc_list = spec_email_df['finance_contact_email'].values.tolist()
    spec_cc = ';'.join(f" {spec}" for spec in spec_cc_list)
    
    # Creating and sending the email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    
    mail.To = finance_emails  # Change this if you need to send to multiple recipients
    mail.Cc = spec_cc
    mail.Subject = 'Upcoming Financial Dues'
    mail.HTMLBody = f"""
    <html>
    <body>
    <p>Hello Finance Team,<br><br>Please find below the list of upcoming financial dues:<br><br>
    {finance_due_table}<br><br>
    Please let us know if you have any questions or concerns.<br><br>
    Thank you,<br>
    Finance Team</p>
    </html>
    </body>
    """
    
    mail.Send()

# Call the function to send the emails
send_finance_due_emails()
