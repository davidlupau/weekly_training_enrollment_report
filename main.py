from functions import load_excel_file, clean_data, remove_duplicates, format_dates, create_report

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

    print("\nSaving facilitator report...")
    report_df = create_report(final_df)
    print("\nReport successfully created.\n")
    
    return report_df
    
if __name__ == "__main__":
    main()