# download_button.py
import streamlit as st
import os

# prog_bar_obj = st.progress(0)

def show_download_button():
    file_path = os.path.join(os.getcwd(), "combined.xlsx")
    file_name = "combined.xlsx"

    if os.path.exists(file_path):
        # Read the file content as bytes
        with open(file_path, "rb") as f:
            file_bytes = f.read()

        st.download_button(
            label="Download Combined Excel File",
            data=file_bytes,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        # Optional: remove after download (commented out for control)
        os.remove(file_path)