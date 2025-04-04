import pandas as pd
import os
import re

def load_excel_file(file_path):
    """
    Loads an Excel file and returns a pandas DataFrame.
    Think of this as your 'mise en place' - getting your ingredients ready.
    
    Args:
        file_path (str): Path to your Excel file
        
    Returns:
        pd.DataFrame: Loaded data in a pandas DataFrame
    """
    try:
        return pd.read_excel(file_path, header=0)
    except Exception as e:
        print(f"Error loading file: {e}")
        return None
    
def clean_data(df):
   """
   Main data cleaning function that handles all transformations in the correct order.
   """
   if df is None:
       return None
   
   # Create a copy to work with
   cleaned_df = df.copy()
   
   # Set proper column names from row 1
   cleaned_df.columns = cleaned_df.iloc[0]
   
   # Remove the first two rows
   print("\nRemoving first rows...")
   cleaned_df = cleaned_df.iloc[2:]
   cleaned_df = cleaned_df.reset_index(drop=True)

   # Convert offering_start_date to datetime format
   print("\nConverting dates to datetime...")
   cleaned_df['Offering Start Date'] = pd.to_datetime(
       cleaned_df['Offering Start Date'], 
       format='%d/%m/%Y %H:%M'
   )
   
   # Remove specified columns
   columns_to_remove = [
       "Blended Course", "Language", "Offering Number", "Topic",
       "Worker Type", "SUB SBU", "Expiration Date",
       "Record Progress Percentage", "Due Date", "Completion Date",
       "Drop Date", "Drop Reason", "Drop Reason - Reference ID",
       "Latest Registration", "Training Contact", "Offering End Date",
       "Allowed Instructors", "Course Grade", "Description"
   ]
   print("\nRemoving non-essential columns...")
   cleaned_df = cleaned_df.drop(columns=columns_to_remove)

   # Rename columns to Python standard
   column_mapping = {
       "Training Title": "training_title",
       "PPG ID": "ppg_id",
       "Full Name": "full_name",
       "Employee Email": "employee_email",
       "Manager": "manager",
       "Manager's Email": "manager_email",
       "Work Country": "work_country",
       "Location": "location",
       "SBU": "sbu",
       "Job Function": "job_function",
       "Registration Date": "registration_date",
       "Registration Status": "registration_status",
       "Completion Status": "completion_status",
       "Offering Start Date": "offering_start_date",
       "Instructors": "instructors",
       "Primary Location": "primary_location",
       "Course Attendance Status": "attendance_status"
   }
   cleaned_df = cleaned_df.rename(columns=column_mapping)

   # Convert ppg_id to numeric
   print("\nConverting PPG IDs to numeric format...")
   non_numeric = cleaned_df[pd.to_numeric(cleaned_df['ppg_id'], errors='coerce').isna()]
   if len(non_numeric) > 0:
       print("Warning: Found non-numeric PPG IDs:")
       print(non_numeric['ppg_id'].unique())
   
   cleaned_df['ppg_id'] = pd.to_numeric(cleaned_df['ppg_id'], errors='coerce')
   if cleaned_df['ppg_id'].isna().sum() == 0:
       cleaned_df['ppg_id'] = cleaned_df['ppg_id'].astype(int)
   
   # Update registration status
   print("\nUpdating registration status values...")
   cleaned_df['registration_status'] = cleaned_df['registration_status'].replace(
       'Enrolled - Pending Approval', 'Pending Approval'
   )

   # Update country names
   cleaned_df['work_country'] = cleaned_df['work_country'].replace({
    'United States of America': 'USA',
    'United Kingdom': 'UK'
    })

   print("\nBasic data cleaning completed.")
   return cleaned_df

def remove_duplicates(df):
   """
   Removes duplicate entries based on business rules for attendance status.
   Must be called after clean_data and before format_dates.
   """
   print("\nStarting duplicate removal process...")
   original_count = len(df)
   
   def select_priority_row(group):
       if len(group) == 1:
           return group
       
       # Check for 'Attended' rows
       attended_rows = group[group['attendance_status'] == 'Attended']
       if not attended_rows.empty:
           return attended_rows.sort_values('offering_start_date', ascending=False).head(1)
       
       # Check for 'Not Entered' rows
       not_entered_rows = group[group['attendance_status'] == 'Not Entered']
       if not not_entered_rows.empty:
           return not_entered_rows.sort_values('offering_start_date', ascending=False).head(1)
       
       # For 'Did Not Attend' rows, keep most recent
       return group.sort_values('offering_start_date', ascending=False).head(1)
   
   df_no_duplicates = df.groupby(['ppg_id', 'training_title'], as_index=False).apply(select_priority_row)
   
   if isinstance(df_no_duplicates.index, pd.MultiIndex):
       df_no_duplicates = df_no_duplicates.reset_index(drop=True)
   
   rows_removed = original_count - len(df_no_duplicates)
   print(f"Rows removed: {rows_removed}")
   
   return df_no_duplicates

def format_dates(df):
   """
   Formats dates to 'Month Day_with_ordinal' format.
   Must be called after remove_duplicates as final step.
   """
   def add_ordinal(day):
       if 10 <= day % 100 <= 20:
           suffix = 'th'
       else:
           suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
       return f"{day}{suffix}"
   
   df['offering_start_date'] = df['offering_start_date'].apply(
       lambda x: f"{x.strftime('%B')} {add_ordinal(x.day)}"
   )
   return df

def clean_name(name):
    """
    Removes content in parentheses from names.
    """
    # Remove both regular and Chinese parentheses and their content
    cleaned = re.sub(r'\s*[\(（].*?[\)）]\s*', '', name)
    return cleaned.strip()

def create_report(df):
    """
    Creates enrollment report with one row per employee showing all module details.
    """
    # Module name mapping
    module_map = {
        'EoL Module 1 - Communication: Connect Through Conversation (Multi-Lingual)': 'Module 1',
        'EoL Module 2 - Coaching: Move People Forward (Multi-Lingual)': 'Module 2',
        'EoL Module 3 - Resolving Workplace Conflict - (Multi-Lingual)': 'Module 3',
        'EoL Module 4 - Delegating: Engage and Empower People (Multi-Lingual)': 'Module 4',
        'EoL Module 5 - Executing Strategy at the Front Line (Multi-Lingual)': 'Module 5',
        'EoL Module 6 - Driving Change (Multi-Lingual)': 'Module 6',
        'Reconnect Day - EOL (Multi-Lingual)': 'Reconnect day'
    }
    
    # Create empty DataFrame for results
    employees = df['ppg_id'].unique()
    columns = ['PPG ID', 'Name', 'Email', 'Location', 'Country', "SBU"]
    
    for module in module_map.values():
        columns.extend([
            f'{module} - Status',
            f'{module} - Attendance',
            f'{module} - Date',
            f'{module} - Facilitator'
        ])
    
    report_df = pd.DataFrame(columns=columns)
    
    # Fill data for each employee
    for ppg_id in employees:
        employee_data = df[df['ppg_id'] == ppg_id].iloc[0]
        row = {
            'PPG ID': ppg_id,
            'Name': clean_name(employee_data['full_name']),
            'Email': employee_data['employee_email'],
            'Country': employee_data['work_country'],
            'Location': employee_data['location'],
            'SBU': employee_data['sbu']

        }
        
        # Get module data
        employee_modules = df[df['ppg_id'] == ppg_id]
        for old_name, new_name in module_map.items():
            module_data = employee_modules[employee_modules['training_title'] == old_name]
            if not module_data.empty:
                data = module_data.iloc[0]
                row.update({
                    f'{new_name} - Status': data['registration_status'],
                    f'{new_name} - Attendance': data['attendance_status'],
                    f'{new_name} - Date': data['offering_start_date'],
                    f'{new_name} - Facilitator': data['instructors']
                })
        
        report_df = pd.concat([report_df, pd.DataFrame([row])], ignore_index=True)
    
    # Create outputs directory if it doesn't exist
    os.makedirs('outputs', exist_ok=True)

    # Save report with date in filename
    today = pd.Timestamp.now().strftime('%Y%m%d')
    filename = os.path.join('outputs', f'{today}_enrollment_report.xlsx')
    report_df.to_excel(filename, index=False)
    
    return report_df

def process_manager_emails(df):
    """
    Process manager emails and update recipient lists.
    
    Args:
        df (pd.DataFrame): The main enrollment DataFrame
    """
    print("\nProcessing manager email addresses...")
    
    # Get unique manager emails from the main DataFrame
    manager_emails = df['manager_email'].unique()
    
    try:
        # Read existing recipients
        existing_recipients = pd.read_excel('manager_guide_recipients.xlsx')
        existing_emails = existing_recipients.iloc[:, 0].tolist()  # Take first column
    except FileNotFoundError:
        print("No existing recipients file found. Creating new one.")
        existing_emails = []
    
    # Find new email addresses
    new_emails = [email for email in manager_emails if email not in existing_emails]
    
    if new_emails:
        print(f"Found {len(new_emails)} new manager email(s)")
        
        # Save new emails to dated file
        today = pd.Timestamp.now().strftime('%Y%m%d')
        new_emails_df = pd.DataFrame(new_emails, columns=['Email'])
        new_file = os.path.join('outputs', f'{today}_managers.xlsx')
        new_emails_df.to_excel(new_file, index=False)
        print(f"New manager emails saved to {new_file}")
        
        # Update master recipient list
        all_emails = existing_emails + new_emails
        updated_df = pd.DataFrame(all_emails, columns=['Email'])
        updated_df.to_excel('manager_guide_recipients.xlsx', index=False)
        print("Master recipient list updated")
    else:
        print("No new manager emails found")

def create_session_report(df):
    """
    Creates a report of all sessions with enrollment counts.
    Groups by module, date, and facilitator.
    """
    print("\nCreating session summary report...")
    
    # Create session summary
    session_summary = df.groupby(
        ['training_title', 'offering_start_date', 'instructors']
    ).size().reset_index(name='enrolled_count')
    
    # Rename training_title to shorter module names
    module_map = {
        'EoL Module 1 - Communication: Connect Through Conversation (Multi-Lingual)': 'Module 1',
        'EoL Module 2 - Coaching: Move People Forward (Multi-Lingual)': 'Module 2',
        'EoL Module 3 - Resolving Workplace Conflict - (Multi-Lingual)': 'Module 3',
        'EoL Module 4 - Delegating: Engage and Empower People (Multi-Lingual)': 'Module 4',
        'EoL Module 5 - Executing Strategy at the Front Line (Multi-Lingual)': 'Module 5',
        'EoL Module 6 - Driving Change (Multi-Lingual)': 'Module 6',
        'Reconnect Day - EOL (Multi-Lingual)': 'Reconnect day'
    }
    
    session_summary['module'] = session_summary['training_title'].map(module_map)
    
    # Reorder and rename columns
    session_summary = session_summary[[
        'module',
        'offering_start_date',
        'instructors',
        'enrolled_count'
    ]].rename(columns={
        'offering_start_date': 'date',
        'instructors': 'facilitator',
        'enrolled_count': 'number_of_employees'
    })
    
    # Sort by date and module
    session_summary = session_summary.sort_values(['date', 'module'])
    
    # Save to Excel
    today = pd.Timestamp.now().strftime('%Y%m%d')
    filename = os.path.join('outputs', f'{today}_session_summary.xlsx')
    session_summary.to_excel(filename, index=False)
    
    print(f"Session summary saved to {filename}")
    
    return session_summary
