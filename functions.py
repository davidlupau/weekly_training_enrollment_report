import pandas as pd

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
    Cleans the data by removing unnecessary rows and columns.
    
    Args:
        df (pd.DataFrame): Input DataFrame to clean
        
    Returns:
        pd.DataFrame: Cleaned DataFrame
    """
    if df is None:
        return None
    
    # Create a copy to work with
    cleaned_df = df.copy()
    
    # First, set proper column names from row 1 (index 0)
    cleaned_df.columns = cleaned_df.iloc[0]
    
    # Now remove the first two rows (the title row and the header row we just used)
    cleaned_df = cleaned_df.iloc[2:]
    
    # Reset the index after dropping rows
    cleaned_df = cleaned_df.reset_index(drop=True)
    
    # Define columns to remove
    columns_to_remove = [
        "Blended Course",
        "Language",
        "Offering Number",
        "Topic",
        "Worker Type",
        "SUB SBU",
        "Expiration Date",
        "Record Progress Percentage",
        "Due Date",
        "Completion Date",
        "Drop Date",
        "Drop Reason",
        "Drop Reason - Reference ID",
        "Latest Registration",
        "Training Contact",
        "Offering End Date",
        "Allowed Instructors",
        "Course Grade",
        "Description"
    ]
    
    # Remove specified columns
    cleaned_df = cleaned_df.drop(columns=columns_to_remove)
    
    # Print information about what we did
    print("\nCleaning summary:")
    print(f"Columns in cleaned data: {len(cleaned_df.columns)}")
    print("\nRemaining columns:")
    for col in cleaned_df.columns:
        print(f"- {col}")
    
    return cleaned_df