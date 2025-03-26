import streamlit as st
import pandas as pd
import io

st.title("Non-Registrant Checker (Excel Files with Emails)")

# Upload the Excel files
registrant_file = st.file_uploader("Upload Registrant List (Excel)", type=['xlsx'])
other_file = st.file_uploader("Upload Attendee/Non-Registrant List (Excel)", type=['xlsx'])

if registrant_file and other_file:
    # Read the files into DataFrames
    registrants_df = pd.read_excel(registrant_file)
    other_df = pd.read_excel(other_file)

    # Preview uploaded data
    st.subheader("Registrant File Preview")
    st.write(registrants_df.head())

    st.subheader("Other List Preview")
    st.write(other_df.head())

    # Make sure they have an 'Email' column (adjust column name as needed)
    if 'Email' in registrants_df.columns and 'Email' in other_df.columns:
        # Clean up email columns (remove NaN, spaces, and make lowercase for matching)
        registrant_emails = registrants_df['Email'].dropna().str.strip().str.lower().tolist()
        other_emails = other_df['Email'].dropna().str.strip().str.lower().tolist()

        # Find non-registrants: emails in other_list not in registrants
        non_registrant_emails = [email for email in other_emails if email not in registrant_emails]

        st.subheader("Non-Registrant Emails Found")
        if non_registrant_emails:
            st.write(non_registrant_emails)

            # Optional: Return more data (names + emails)
            non_registrants_full = other_df[other_df['Email'].str.lower().isin(non_registrant_emails)]

            # Create a BytesIO buffer to hold the Excel data
            output = io.BytesIO()

            # Write to the Excel file in memory
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                non_registrants_full.to_excel(writer, index=False, sheet_name='Non-Registrants')

            # Get the Excel data from the buffer
            processed_data = output.getvalue()

            # Download button
            st.download_button(
                label="Download Non-Registrant List (Excel)",
                data=processed_data,
                file_name='non_registrants.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.write("All attendees are registered!")
    else:
        st.error("The required 'Email' column was not found in one or both files. Please check your files.")
