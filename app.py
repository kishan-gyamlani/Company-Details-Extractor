import streamlit as st
import pandas as pd

def search_company_in_excel(uploaded_file, company_name):
    """
    Searches for a company name in all sheets of an uploaded Excel file.

    Args:
    - uploaded_file (UploadedFile): Uploaded Excel file from Streamlit
    - company_name (str): Company name to search for.

    Returns:
    - results (dict): Dictionary of search results with sheet names as keys
    """
    # Read the uploaded Excel file
    excel_data = pd.ExcelFile(uploaded_file)
    results = {}

    for sheet_name in excel_data.sheet_names:
        sheet_df = excel_data.parse(sheet_name)
        
        if 'Company' in sheet_df.columns:
            # Convert Company column to string to ensure string operations work
            sheet_df['Company'] = sheet_df['Company'].astype(str)

            # Search for company name (case-insensitive)
            matching_rows = sheet_df[sheet_df['Company'].str.contains(company_name, case=False, na=False)]

            if not matching_rows.empty:
                results[sheet_name] = matching_rows

    return results

def main():
    st.title("Company Search in Excel")
    
    # Upload Excel file
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    
    if uploaded_file is not None:
        # Input for company name
        company_name = st.text_input("Enter Company Name")

        if company_name:
            # Call the search function
            search_results = search_company_in_excel(uploaded_file, company_name)

            if search_results:
                st.success(f"Found {sum(len(df) for df in search_results.values())} matching result(s):")
                
                for sheet_name, matching_df in search_results.items():
                    st.subheader(f"Results in Sheet: {sheet_name}")
                    
                    # Display the matching rows as a table
                    st.dataframe(matching_df)
                    
                    st.write("-" * 50)
            else:
                st.warning("No matches found.")

if __name__ == "__main__":
    main()