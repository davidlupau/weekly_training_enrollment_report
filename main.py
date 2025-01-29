from functions import load_excel_file, clean_data

def main():
    # Load the data
    print("Loading data...")
    input_df = load_excel_file('input_data.xlsx')
    
    # Clean the data
    print("\nCleaning data...")
    cleaned_df = clean_data(input_df)
    
    # Save the cleaned data
    print("\nSaving data...")
    cleaned_df.to_excel('cleaned_data.xlsx', index=False)
    
if __name__ == "__main__":
    main()