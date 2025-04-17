from combineExcel import getFormattedSheet
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import os
import json
import re
import zipfile
from utils import show_download_button
UPLOAD_FOLDER = "files"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# Initialize session state
if 'processed' not in st.session_state:
    st.session_state.processed = False

st.title("Excel Automation")

directory_path = os.path.join(os.getcwd(), UPLOAD_FOLDER)
xlsx_files = [f for f in os.listdir(directory_path) if f.endswith(".xlsx")]
if len(xlsx_files) > 0:
    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"Error deleting {file_path}: {e}")
uploaded_file = st.file_uploader("Upload files", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_file and not st.session_state.processed:
    for file in uploaded_file:
        file_path = os.path.join(directory_path, file.name)
        with open(file_path, "wb") as f:
            f.write(file.getbuffer())

    xlsx_files = [f for f in os.listdir(directory_path) if f.endswith(".xlsx")]

    actualheader = []
    header = []
    if len(xlsx_files) > 0 and uploaded_file:
        success_msg = st.success(f"Files uploaded successfully")
        workbook = []
        # from utils import prog_bar_obj
        prog_bar_obj = st.progress(0, "Processing files. Please wait...")
        for file in xlsx_files:
            try:
                workbook.append(pd.ExcelFile(os.path.join(directory_path, file), engine='openpyxl'))
            except zipfile.BadZipFile:
                print(f"Skipping invalid or corrupted file: {file}")
            except Exception as e:
                print(f"Error reading {file}: {e}")

        sheet_name = workbook[0].sheet_names[2:len(workbook[0].sheet_names)-1]
        for i, sheet in enumerate(sheet_name):
            progress = int((i / len(sheet_name)) * 100)
            prog_bar_obj.progress(progress, f"sheet: {sheet}")
            combainedJson = {
                "sheet": sheet,
                "title": "",
                "page": []
            }
            pdsheet = pd.read_excel(workbook[0], sheet)
            cleaned_data = pdsheet.dropna(how='all').dropna(axis=1, how='all')
            page_name_array = []
            page_no = 0
            for index, row in cleaned_data.iterrows():
                page_n = row.dropna().values[0]
                if isinstance(page_n, str):
                    if "PAGE" in page_n and page_n not in page_name_array and len(row.dropna()) == 1:
                        page_name_array.append(page_n)
                        if page_n[-1] == "S":
                            break

            title1 = cleaned_data.loc[:, ~cleaned_data.columns.str.contains('^Unnamed')].columns[0]
            title2 = cleaned_data.iloc[0].dropna().values[0]
            title3 = cleaned_data.iloc[1].dropna().values[0]
            title4 = cleaned_data.iloc[2].dropna().values[0]
            title5 = cleaned_data.iloc[3].dropna().values[0]
            combainedJson["title"] = f"{title1}\n{title2}\n{title3}\n{title4}\n{title5}"
            startInx = [0 for i in range(len(workbook))]

            for page in page_name_array:
                success_msg.empty()
                prog_bar_obj.progress(progress, f"Sheet: {sheet} - Page: {page}")
                header = []
                page_info_item = {
                    "page_name": "",
                    "Contractor": [],
                }
                page_info_item["page_name"] = page

                for contractor in workbook:
                    sheetdata = pd.read_excel(contractor, sheet)
                    preprocessed_data = sheetdata.dropna(how='all').dropna(axis=1, how='all')
                    preprocessed_data.reset_index(drop=True, inplace=True)

                    for index, row in preprocessed_data.iloc[startInx[workbook.index(contractor)]:].iterrows():
                        x = row.dropna().values.tolist()
                        if all(str(item).replace(' ', '').isalpha() for item in x) and len(header) == 0 and row.isnull().sum() <= 1:
                            for item in x:
                                header.append(item)

                        if (set(row.values) & set(header)) and preprocessed_data.iloc[index+1].dropna().values[0] == page:
                            Contractor_page_info_item = {
                                "contractor_name": "",
                                "bill_section_header": [],
                                "bill_section_items": [],
                                "total": ""
                            }
                            Contractor_page_info_item["contractor_name"] = os.path.splitext(os.path.split(contractor)[1])[0]

                            if pd.isna(row.values[5]):
                                preprocessed_data.at[index, preprocessed_data.columns[5]] = 'AMOUNT SAR RIYAL'
                                if 'AMOUNT SAR RIYAL' not in header:
                                    header.append('AMOUNT SAR RIYAL')
                                inx = index + 1
                                non_billable_description = ""
                                item_array = []

                            while not (set(preprocessed_data.iloc[inx].values) & set(header)) and inx < len(preprocessed_data) - 1:
                                header_data = str(preprocessed_data.iloc[inx].dropna().values[0])

                                def bill_section():
                                    page_n_total_not_in_data = "PAGE" not in header_data and "Total" not in header_data
                                    if pd.isna(preprocessed_data.iloc[inx].values[0]) and page_n_total_not_in_data:
                                        global non_billable_description
                                        non_billable_description = f"{non_billable_description}\n{preprocessed_data.iloc[inx].dropna().values[0]}"

                                bill_section()

                                if re.fullmatch(r"[A-Z]", str(preprocessed_data.iloc[inx].values[0])):
                                    Contractor_page_info_item["bill_section_header"].append(non_billable_description)
                                    non_billable_description = ""
                                    while re.fullmatch(r"[A-Z]", str(preprocessed_data.iloc[inx].values[0])):
                                        item_array.append({
                                            f"{header[0]}": preprocessed_data.iloc[inx].values[0],
                                            f"{header[1]}": preprocessed_data.iloc[inx].values[1],
                                            f"{header[2]}": preprocessed_data.iloc[inx].values[2],
                                            f"{header[3]}": preprocessed_data.iloc[inx].values[3],
                                            f"{header[4]}": preprocessed_data.iloc[inx].values[4],
                                            f"{header[5]}": preprocessed_data.iloc[inx].values[5]
                                        })
                                        inx = inx + 1
                                    Contractor_page_info_item["bill_section_items"].append(item_array)
                                    item_array = []
                                    bill_section()
                                elif "PAGE" in header_data and "PAGE" in str(preprocessed_data.iloc[inx].values[1]):
                                    Contractor_page_info_item["bill_section_header"].append(non_billable_description)
                                    non_billable_description = ""
                                    while "Total" not in str(preprocessed_data.iloc[inx].dropna().values[0]):
                                        item_array.append({
                                            f"{header[0]}": preprocessed_data.iloc[inx].values[0],
                                            f"{header[1]}": preprocessed_data.iloc[inx].values[1],
                                            f"{header[2]}": preprocessed_data.iloc[inx].values[2],
                                            f"{header[3]}": preprocessed_data.iloc[inx].values[3],
                                            f"{header[4]}": preprocessed_data.iloc[inx].values[4],
                                            f"{header[5]}": preprocessed_data.iloc[inx].values[5]
                                        })
                                        inx = inx + 1
                                    Contractor_page_info_item["bill_section_items"].append(item_array)
                                    item_array = []

                                if pd.isna(preprocessed_data.iloc[inx].values).any() and "Total" in str(preprocessed_data.iloc[inx].dropna().values[0]):
                                    while not (set(preprocessed_data.iloc[inx].values) & set(header)) and inx < len(preprocessed_data) - 1:
                                        total = preprocessed_data.iloc[inx].dropna().values[0]
                                        Contractor_page_info_item["total"] = 0
                                        if isinstance(total, (float, int)):
                                            Contractor_page_info_item["total"] = total
                                            page_info_item["Contractor"].append(Contractor_page_info_item)
                                            break
                                        inx = inx + 1

                                    if Contractor_page_info_item not in page_info_item["Contractor"]:
                                        page_info_item["Contractor"].append(Contractor_page_info_item)
                                    break
                                inx = inx + 1
                        else:
                            continue
                        startInx[workbook.index(contractor)] = inx + 1
                        break
                combainedJson["page"].append(page_info_item)
            work_book = getFormattedSheet(json.dumps(combainedJson), sheet_name.index(sheet), sheet, progress)

        prog_bar_obj.progress(100, "Completed!")
        success_msg.success("Excel file created successfully!")

        # Prevent rerun
        st.session_state.processed = True
        # Show download button
        show_download_button()
else:
    print("no files found")

