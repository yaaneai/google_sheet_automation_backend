# from sampleJson import formattedSheet
import json
from openpyxl import Workbook
from openpyxl.styles import Alignment, PatternFill, Border, Font

def increment_column(column):
    # If column is Z, return AA
    if column == 'Z':
        return 'AA'
    # If column has multiple characters
    if len(column) == 1:
        next_char = chr(ord(column) + 1)
        return next_char
    else:
        last_char = column[-1]
        if last_char == 'Z':
            return increment_column(column[:-1]) + 'A'
        else:
            return column[:-1] + chr(ord(last_char) + 1)


def getFormattedSheet(sheetJson):
    formatted_josn = json.loads(sheetJson)    
    header = list(formatted_josn["page"][0]["Contractor"][0]["bill_section_items"][0][0].keys())
    number_of_contractors = len(formatted_josn["page"][0]["Contractor"])
    commonColumns = header[0:len(header)-2]
    contractorColums = header[len(header)-2:len(header)]
    header.extend(header[len(header)-2:len(header)]*(number_of_contractors-1))
    startColumn = 'A'
    endColumn = ''
    Row_num_to_insert = 1
    for i in range(1,len(header)):
        endColumn = increment_column(startColumn if i==1 else endColumn)
    wb = Workbook()
    sheet = wb.create_sheet(formatted_josn["sheet"])
    title = str.splitlines(formatted_josn["title"])
    # add titles in the new sheet 
    for i in range(1,len(title)):
        sheet.merge_cells(f"{startColumn}{i}:{endColumn}{i}")
        sheet[f"{startColumn}{i}"]= title[i]
        sheet[f"{startColumn}{i}"].alignment = Alignment(horizontal="center", vertical="center")
        Row_num_to_insert = Row_num_to_insert+1
    #creates merged cell contractor name header column
    contracter_merge_cells_start_column = chr(ord(startColumn) + len(commonColumns))
    contracter_merge_cells_end_column = ''
    for i in range(1,number_of_contractors+1):
        def contractName():
            startCell =f"{contracter_merge_cells_start_column}{Row_num_to_insert}"
            sheet.merge_cells(f"{startCell}:{contracter_merge_cells_end_column}{Row_num_to_insert}")
            sheet[startCell]= formatted_josn["page"][0]["Contractor"][i-1]["contractor_name"]
            sheet[startCell].alignment = Alignment(horizontal="center", vertical="center")
            sheet[startCell].font = Font(bold=True)
        if i == 1:
            contracter_merge_cells_end_column = chr(ord(contracter_merge_cells_start_column) + 
                                                    len(contractorColums)-1)
            contractName()
        else:
            contracter_merge_cells_start_column = chr(ord(contracter_merge_cells_end_column) + 1)
            contracter_merge_cells_end_column = chr(ord(contracter_merge_cells_start_column) + len(contractorColums)-1)
            contractName() 
    Row_num_to_insert=Row_num_to_insert+1
    # getting into each page info
    for page in formatted_josn["page"]:
        # add header to page:
        #page header column names
        header_column = {
            f"{header[0]}":"",
            f"{header[1]}": "",
            f"{header[2]}": "",
            f"{header[3]}": "",
            f"{header[4]}": [],
            f"{header[5]}": []
        }
        for i in range(1,len(header)+1):
            column_to_insert = f"{chr(ord(startColumn) + (i-1))}" 
            cell_to_insert = f"{column_to_insert}{Row_num_to_insert}"
            sheet[cell_to_insert]=header[i-1]
            sheet[cell_to_insert].font = Font(bold=True)
            sheet[cell_to_insert].alignment = Alignment(horizontal="center", vertical="center")
            sheet[cell_to_insert].fill = PatternFill(start_color="A6A6A6", end_color="A6A6A6", fill_type="solid")
            if header[i-1] in commonColumns:
                header_column[header[i-1]]= column_to_insert
                
            else:
                header_column[header[i-1]].append(column_to_insert)
        Row_num_to_insert=Row_num_to_insert+1
        #page name under "AMOUNT SAR RIYAL" row
        for i in range(1,number_of_contractors+1):
            cell_to_insert = f"{chr(ord(startColumn) + (len(commonColumns)+len(contractorColums*i)-1))}{Row_num_to_insert}"
            sheet[cell_to_insert]=page["page_name"]
            sheet[cell_to_insert].font = Font(bold=True)
            sheet[cell_to_insert].alignment = Alignment(horizontal="center", vertical="center")
        Row_num_to_insert=Row_num_to_insert+1
        #section_header and section_item
        for section_header_counter in range(0,len(page["Contractor"][0]["bill_section_header"])):
            section_header_split_list = str.splitlines(page["Contractor"][0]["bill_section_header"][section_header_counter])
            for section_header_line in section_header_split_list:
                if page["page_name"] not in section_header_line and len(section_header_line) != 0:
                    cell_to_insert = f"{header_column[f"{header[1]}"]}{Row_num_to_insert}"
                    sheet[cell_to_insert]= section_header_line
                    Row_num_to_insert=Row_num_to_insert+1
                else:
                    continue
            Row_num_to_insert=Row_num_to_insert+1
            for section_item_counter in range(0,len(page["Contractor"][0]["bill_section_items"][section_header_counter])):
                combained_section_item = {
                    f"{header[0]}":"",
                    f"{header[1]}": "",
                    f"{header[2]}": 0,
                    f"{header[3]}": "",
                    f"{header[4]}": [],
                    f"{header[5]}": []
                }
                # get the items from each contractor combine and store it as dict 
                for contractor in range(0, number_of_contractors):
                    print(page["page_name"])
                    print(page["Contractor"][contractor]["contractor_name"])
                    print(page["Contractor"][contractor]["bill_section_items"][section_header_counter])
                    print(page["Contractor"][contractor]["bill_section_items"][section_header_counter][section_item_counter])

                    item = page["Contractor"][contractor]["bill_section_items"][section_header_counter][section_item_counter]
                    if contractor == 0:
                        combained_section_item[f"{header[0]}"]=item[f"{header[0]}"]
                        combained_section_item[f"{header[1]}"]=item[f"{header[1]}"]
                        combained_section_item[f"{header[2]}"]=item[f"{header[2]}"]
                        combained_section_item[f"{header[3]}"]=item[f"{header[3]}"]
                        combained_section_item[f"{header[4]}"].append(item[f"{header[4]}"])
                        combained_section_item[f"{header[5]}"].append(item[f"{header[5]}"])
                    else:
                        combained_section_item[f"{header[4]}"].append(item[f"{header[4]}"])
                        combained_section_item[f"{header[5]}"].append(item[f"{header[4]}"])
                
                for key, value in combained_section_item.items():
                    if key in commonColumns:
                        cell_to_insert = f"{header_column[key]}{Row_num_to_insert}"
                        sheet[cell_to_insert]=value
                        if key == f"{header[3]}":
                            sheet[cell_to_insert].alignment = Alignment(horizontal="center", vertical="center")
                    else:
                        for i in range(0,len(value)):
                            cell_to_insert = f"{header_column[key][i]}{Row_num_to_insert}"
                            sheet[cell_to_insert]=value[i]
                            if key == f"{header[4]}" and value[i]==max(value):
                                sheet[cell_to_insert].fill = PatternFill(start_color="ff0000", end_color="ff0000", fill_type="solid")
                Row_num_to_insert=Row_num_to_insert+2
                
        for contractor in range(0, number_of_contractors):
            if contractor == 0:
                startCell = f"{startColumn}{Row_num_to_insert}"
                endCell = f"{chr(ord(startColumn) + len(commonColumns)-1)}{Row_num_to_insert}"
                sheet.merge_cells(f"{startCell}:{endCell}")
                sheet[startCell]="Total - Saudi Riyal Carried to Collection"
                sheet[startCell].alignment = Alignment(horizontal="right", vertical="center")
                sheet[startCell].font = Font(bold=True)
            cell_to_insert= f"{header_column[f"{header[5]}"][contractor]}{Row_num_to_insert}"
            total= page["Contractor"][contractor]["total"]
            sheet[cell_to_insert]= total
            sheet[cell_to_insert].font = Font(bold=True)
        Row_num_to_insert=Row_num_to_insert+1
    del wb["Sheet"]
    wb.save("combined.xlsx")

    
# getFormattedSheet(formattedSheet)