from combainedExcel import getFormattedSheet
import pandas as pd
import os
import json
import re
import zipfile
directory_path = "C:\\Users\\Vignesh\\Documents\\google_sheet_automation_backend\\files"
xlsx_files = [f for f in os.listdir(directory_path) if f.endswith(".xlsx")]
actualheader = []
header = []
if len(xlsx_files) > 0:
    workbook = []
    for file in xlsx_files:
        try:
            workbook.append(pd.ExcelFile(f'{directory_path}\\{file}', engine='openpyxl'))
        except zipfile.BadZipFile:
            print(f"Skipping invalid or corrupted file: {file}")
        except Exception as e:
            print(f"Error reading {file}: {e}")
    sheet_name = workbook[0].sheet_names[2:len(workbook[0].sheet_names)-1]#get the list of sheet names from the uploaded contrancter file"
    for sheet in sheet_name:
        print("sheet: ",sheet)
        combainedJson={
        "sheet":sheet,
        "title":"",
        "page":[]
        }
        pdsheet=pd.read_excel(workbook[0],sheet)
        cleaned_data = pdsheet.dropna(how='all').dropna(axis=1, how='all')
        #getting available Page names
        page_name_array = []
        page_no = 0
        for index, row in cleaned_data.iterrows():
            page_n = row.dropna().values[0]
            if isinstance(page_n,str):
                if "PAGE" in page_n and page_n not in page_name_array:
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
            if page[-1] == "S":
                print("debug for Page S")
            print("fetching: ",page)
            header = []
            page_info_item = {
                "page_name":"",
                "Contractor":[],
                }
            #goto_page = page_name_array.index(page)+1
            page_info_item["page_name"] = page
            for contractor in workbook:
                sheetdata=pd.read_excel(contractor,sheet)
                preprocessed_data = sheetdata.dropna(how='all').dropna(axis=1, how='all')
                preprocessed_data.reset_index(drop=True, inplace=True)
                for index, row in preprocessed_data.iloc[startInx[workbook.index(contractor)]:].iterrows():
                    x = row.dropna().values.tolist()
                    #check and collects the header names 
                    if all(str(item).replace(' ', '').isalpha() for item in x) and len(header) == 0 and row.isnull().sum() <= 1:
                        for item in x:
                            header.append(item)
                    if (set(row.values) & set(header)) and preprocessed_data.iloc[index+1].dropna().values[0] == page:
                        Contractor_page_info_item = {
                            "contractor_name":"",
                            "bill_section_header":[],
                            "bill_section_items":[],
                            "total":""
                            }
                        Contractor_page_info_item["contractor_name"]=os.path.splitext(os.path.split(contractor)[1])[0]
                        if pd.isna(row.values[5]):
                            preprocessed_data.at[index, preprocessed_data.columns[5]] = 'AMOUNT SAR RIYAL'
                            if 'AMOUNT SAR RIYAL' not in header:
                                header.append('AMOUNT SAR RIYAL') 
                            inx=index+1
                            non_billable_description = ""
                            item_array = []
                        #get into quotation data collects section header data
                        while not(set(preprocessed_data.iloc[inx].values) & set(header)) and inx < len(preprocessed_data)-1:
                            def bill_section():
                                if pd.isna(preprocessed_data.iloc[inx].values[0]) and "Total" not in str(preprocessed_data.iloc[inx].dropna().values[0]): #collects only bill_section_header
                                    global non_billable_description
                                    non_billable_description = f"{non_billable_description}\n{preprocessed_data.iloc[inx].dropna().values[0]}"
                            bill_section()
                            #finding the item row under section header by checking the item column has 'A'-'Z'
                            if re.fullmatch(r"[A-Z]", str(preprocessed_data.iloc[inx].values[0])):
                                Contractor_page_info_item["bill_section_header"].append(non_billable_description)
                                non_billable_description = ""
                                while re.fullmatch(r"[A-Z]", str(preprocessed_data.iloc[inx].values[0])):# true only ITEM cell has A-Z value
                                    item_array.append({
                                        f"{header[0]}":preprocessed_data.iloc[inx].values[0],
                                        f"{header[1]}": preprocessed_data.iloc[inx].values[1],
                                        f"{header[2]}": preprocessed_data.iloc[inx].values[2],
                                        f"{header[3]}": preprocessed_data.iloc[inx].values[3],
                                        f"{header[4]}": preprocessed_data.iloc[inx].values[4],
                                        f"{header[5]}": preprocessed_data.iloc[inx].values[5]
                                        })
                                    inx=inx+1
                                Contractor_page_info_item["bill_section_items"].append(item_array)
                                item_array = []
                                bill_section()
                            #collects the data from total field of the page  
                            if pd.isna(preprocessed_data.iloc[inx].values).any() and "Total" in str(preprocessed_data.iloc[inx].dropna().values[0]):
                                while not(set(preprocessed_data.iloc[inx].values) & set(header)) and inx < len(preprocessed_data)-1:
                                    total = preprocessed_data.iloc[inx].dropna().values[0]
                                    Contractor_page_info_item["total"]=0
                                    if isinstance(total,(float,int)):
                                        Contractor_page_info_item["total"]=total
                                        page_info_item["Contractor"].append(Contractor_page_info_item)
                                        break
                                    inx=inx+1
                                if Contractor_page_info_item not in page_info_item["Contractor"]:
                                    page_info_item["Contractor"].append(Contractor_page_info_item)
                                break

                            inx=inx+1
                    else:
                        continue
                    startInx[workbook.index(contractor)]=inx+1
                    break
            combainedJson["page"].append(page_info_item)
        getFormattedSheet(json.dumps(combainedJson),sheet_name.index(sheet))
        
else:
  print("no files found")