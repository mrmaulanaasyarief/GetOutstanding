from openpyxl import load_workbook
import re
from pprint import pprint
from openpyxl.styles import Border, Side
from openpyxl.styles import Alignment
from openpyxl.styles import Font
from openpyxl.workbook.child import INVALID_TITLE_REGEX

def main():
    # opening the source excel file
    filename ="C:\\Users\\mrmau\\Documents\\Cevin\\GetOutstanding\\Final April 2023.xlsx"
    workbook = load_workbook(filename, data_only=True)
    sheet = workbook.worksheets[0]

    # delete all sheets except first sheet
    for title in workbook.sheetnames[1:]:
        del workbook[title]
    
    # set the start and end of row of data
    start = 7
    end = 356

    # loop trough row of data
    for i in range(start,end+1):
        # create new sheet with unit + tenant combined as sheet name
        sheet_name = str(sheet["C"+str(i)].value).lstrip(" ") + " " + str(sheet["B"+str(i)].value).lstrip(" ") 
        sheet_name = re.sub(INVALID_TITLE_REGEX, '_', sheet_name)   

        title = sheet_name if len(sheet_name) <= 31 else sheet_name[:31]
        workbook.create_sheet(title) # no more than 31 char
        print(sheet_name)


        # write unit and name on created sheet
        created_sheet = workbook[title]
        created_sheet.column_dimensions['B'].width = 15
        created_sheet.column_dimensions['C'].width = 15
        created_sheet.column_dimensions['D'].width = 15
        created_sheet.column_dimensions['E'].width = 15
        created_sheet.column_dimensions['F'].width = 15

        created_sheet["B2"] = sheet["C"+str(i)].value # unit
        created_sheet["C2"] = sheet["B"+str(i)].value # tenant
        
        # border style
        thin_border = Border(left=Side(style='thin'), 
                     right=Side(style='thin'), 
                     top=Side(style='thin'), 
                     bottom=Side(style='medium'))

        # set dict for all value per month
        data_dict = {}

        # get utility value per month
        value = "Utility"
        data_dict = get_all_total_value(i, value, data_dict, sheet)
        # set table header
        created_sheet["C3"] = value
        created_sheet["C3"].border = thin_border
        created_sheet["C3"].alignment = Alignment(horizontal='center')
        created_sheet["C3"].font = Font(bold=True)


        # get utility value per month
        value = "Service Charge"
        data_dict = get_all_total_value(i, value, data_dict, sheet)
        # set table header
        created_sheet["D3"] = value
        created_sheet["D3"].border = thin_border
        created_sheet["D3"].alignment = Alignment(horizontal='center')
        created_sheet["D3"].font = Font(bold=True)
        
        # get utility value per month
        value = "Sinking Fund"
        data_dict = get_all_total_value(i, value, data_dict, sheet)
        # set table header
        created_sheet["E3"] = value
        created_sheet["E3"].border = thin_border
        created_sheet["E3"].alignment = Alignment(horizontal='center')
        created_sheet["E3"].font = Font(bold=True)

        # set table header total/bulan
        created_sheet["F3"] = "Total/Month"
        created_sheet["F3"].border = thin_border
        created_sheet["F3"].alignment = Alignment(horizontal='center')
        created_sheet["F3"].font = Font(bold=True)

        begin = 4
        total_ut = 0
        total_sc = 0
        total_sf = 0
        total_total_month = 0
        # Iterating over values
        for month, bill in data_dict.items():
            ut = bill["Utility"] if "Utility" in bill else "0"
            sc = bill["Service Charge"] if "Service Charge" in bill else "0"
            sf = bill["Sinking Fund"] if "Sinking Fund" in bill else "0"

            ut = ut.lstrip(" ").replace("-", "") if isinstance(ut, str) else ut
            sc = sc.lstrip(" ").replace("-", "") if isinstance(sc, str) else sc
            sf = sf.lstrip(" ").replace("-", "") if isinstance(sf, str) else sf

            ut = ut if ut != "" else "0"
            sc = sc if sc != "" else "0"
            sf = sf if sf != "" else "0"

            total_month = int(ut)+int(sc)+int(sf)

            total_ut += int(ut)
            total_sc += int(sc)
            total_sf += int(sf)
            total_total_month += int(total_month)

            if(ut != "0" or sc != "0" or sf != "0"):
                created_sheet["B"+str(begin)] = month
                created_sheet["C"+str(begin)] = ut if ut != "0" else ""
                created_sheet["D"+str(begin)] = sc if ut != "0" else ""
                created_sheet["E"+str(begin)] = sf if ut != "0" else ""
                created_sheet["F"+str(begin)] = total_month
                
                created_sheet["B"+str(begin)].border = thin_border
                created_sheet["C"+str(begin)].border = thin_border
                created_sheet["D"+str(begin)].border = thin_border
                created_sheet["E"+str(begin)].border = thin_border
                created_sheet["F"+str(begin)].border = thin_border

                created_sheet["B"+str(begin)].number_format = 'Rp #,##'
                created_sheet["C"+str(begin)].number_format = 'Rp #,##'
                created_sheet["D"+str(begin)].number_format = 'Rp #,##'
                created_sheet["E"+str(begin)].number_format = 'Rp #,##'
                created_sheet["F"+str(begin)].number_format = 'Rp #,##'

                begin += 1
                
        if total_total_month != 0:
            created_sheet["C"+str(begin)] = total_ut
            created_sheet["D"+str(begin)] = total_sc
            created_sheet["E"+str(begin)] = total_sf
            created_sheet["F"+str(begin)] = total_total_month

            created_sheet["C"+str(begin)].number_format = 'Rp #,##'
            created_sheet["D"+str(begin)].number_format = 'Rp #,##'
            created_sheet["E"+str(begin)].number_format = 'Rp #,##'
            created_sheet["F"+str(begin)].number_format = 'Rp #,##'

            created_sheet["E"+str(begin+1)] = "Total Bill"
            created_sheet["F"+str(begin+1)] = total_ut + total_sc + total_sf

            created_sheet["F"+str(begin+1)].number_format = 'Rp #,##'

            created_sheet["E"+str(begin+1)].font = Font(bold=True)
            created_sheet["F"+str(begin+1)].font = Font(bold=True)

            created_sheet["E"+str(begin+1)].alignment = Alignment(horizontal='right')
        else:
            del workbook[title]
            
    del workbook[workbook.sheetnames[0]]
    workbook.save(filename= "C:\\Users\\mrmau\\Documents\\Cevin\\GetOutstanding\\Outstanding April 2023.xlsx")


def content_checker(sheet, value):
    for row in sheet:
        for cell in row:
            if cell.value == value:
                return cell

def merged_span_check(sheet, cell):
    for merged_cell in sheet.merged_cells.ranges:
        if cell.coordinate in merged_cell:
            return merged_cell

def get_all_total_value(i, value, data_dict, sheet):
    # get range value (ex. "A1:Z1") then splitted by ":"
    cell = content_checker(sheet, value)
    range_values = str(merged_span_check(sheet, cell)).split(":")

    # loop trough the range
    for r in sheet[range_values[0]:range_values[1]]:
        for x in r:
            # split letter and number using regex
            match = re.match(r"([a-z]+)([0-9]+)", x.coordinate, re.I)
            if match:
                # get splitted letter and number as an array
                coordinates = match.groups()


                month = sheet[coordinates[0]+str(int(coordinates[1])+1)].value
                # if the cell value is not empty
                if sheet[coordinates[0]+str(i)].value is not None:
                    total = sheet[coordinates[0]+str(i)].value
                else:
                    total = ""

                if month not in data_dict:
                    data_dict[month] = {}
                
                data_dict[month][value] = total
    return data_dict

if __name__ == '__main__':
    main()