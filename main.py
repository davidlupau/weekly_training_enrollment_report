from functions import load_excel_file, clean_data, remove_duplicates, format_dates, create_report, process_manager_emails

def main():
    # Load the data
    print("Loading data...")
    input_df = load_excel_file('input_data.xlsx')
    
    # Clean and process the data
    print("\nCleaning data...")
    cleaned_df = clean_data(input_df)
    deduped_df = remove_duplicates(cleaned_df)
    final_df = format_dates(deduped_df)
    
    # Save the final data
    print("\nSaving data...")
    final_df.to_excel('cleaned_data.xlsx', index=False)

    print("\ncleaned_data.xlsx successfully created.")

    print("\nSaving enrollment report...")
    report_df = create_report(final_df)
    print("\nEnrollment report successfully created.\n")

    print("\nExtracting managers email address...")
    process_manager_emails(final_df)
    print("\nManagers email addresses successfully extracted.\n")
    
    return report_df
    
if __name__ == "__main__":
    main()