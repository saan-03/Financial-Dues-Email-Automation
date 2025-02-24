# **Financial Dues Email Automation**

## **Objective:**

The primary goal of this project is to automate the process of sending emails containing a list of upcoming financial dues to finance teams or specific vendors. This project extracts data from a database, processes it, and sends out formatted emails with the relevant financial information.

**Technologies Used:**

- **Python**: The core programming language used for the automation.
- **DuckDB**: A lightweight database used for querying financial due data.
- **Pandas**: For handling and processing data frames, and converting them to HTML for email formatting.
- **Win32com Client**: Used to interface with Microsoft Outlook for email sending.
- **Inputs Module**: A custom module (assumed to exist) for accessing external data like vendor emails.

---

## **Features:**

1. **Database Connection:**
    - Connects to a DuckDB database (`connections.db`).
    - Queries financial due data from the `finance_due_data` table, filtering by due dates that are in the future.
2. **Data Processing:**
    - Uses SQL queries to retrieve data on upcoming financial dues, including vendor names, due dates, invoice numbers, and amounts due.
    - This data is processed into a pandas DataFrame for easy manipulation and HTML conversion.
3. **Email Preparation:**
    - Retrieves the contact emails of finance team members or vendor-specific contacts from a pre-existing dataframe (`finance_emails_df`).
    - Queries the database for additional finance-related contact emails (in the case of CCing additional recipients).
    - Formats the financial due data into an HTML table for clean email presentation.
4. **Automated Email Sending:**
    - Uses the `win32com.client` library to interface with Microsoft Outlook and send automated emails.
    - Each email contains the upcoming financial due details formatted as an HTML table.
    - The email recipients (To and CC) are populated based on data from the `finance_emails_df` and the SQL query for finance contacts.

---

## **Process Flow:**

1. **Query Financial Data:**
    - The code fetches the relevant financial data from the database (`finance_due_data`), including upcoming due dates and amounts.
2. **Format Data:**
    - The retrieved data is converted into an HTML table using `pandas.to_html()` for easy integration into the email body.
3. **Send Emails:**
    - Emails are sent to specified finance contact emails using Microsoft Outlook.
    - The HTML body of the email includes a table summarizing the upcoming dues, making it visually appealing and easy to understand for the recipient.
    - CC recipients are added based on the query results, ensuring that the necessary stakeholders are informed.

---

## **Impact and Benefits:**

1. **Increased Efficiency & Time Savings:**
    
    Automating the email notification process eliminates manual tasks, ensuring timely reminders are sent without administrative overhead, allowing teams to focus on higher-priority tasks.
    
2. **Reduced Errors & Improved Accuracy:**
    
    Automation ensures that financial data is consistently accurate, reducing the risk of human error in both data handling and email distribution, ensuring reliable communication.
    
3. **Better Communication & Stakeholder Alignment:**
    
    The automated emails ensure all relevant stakeholders are informed and receive clear, professional communication, improving transparency and collaboration across teams.
