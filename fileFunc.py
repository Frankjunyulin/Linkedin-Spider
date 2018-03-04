import xlrd
import xlwt
import os.path
from xlrd import open_workbook
from xlwt import easyxf
from xlutils.copy import copy

#Help function to work with file

def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
    first_sheet = book.sheet_by_index(0)
    cell = first_sheet.cell(0,0)
    max_row = len(first_sheet.col_values(0))
    result_data =[]
    for curr_row in range(1, max_row, 1):
        row_data = []
        for curr_col in range(0, 3, 1):
            data = first_sheet.cell_value(curr_row, curr_col) # Read the data in the current cell
            row_data.append(data)
        result_data.append(row_data)

    for proName in result_data:
        temp = str(proName[0])
        temp = temp.split('.0')
        proName[0] = temp[0]
    return(result_data)



def write_file(path):
    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
        num_format_str='#,##0.00')
    style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

    wb = xlwt.Workbook()
    ws = wb.add_sheet('summary')

    ws.write(0, 0, 'our_id')
    ws.write(0, 1, 'our_firstname')
    ws.write(0, 2, 'our_name')
    ws.write(0, 3, 'id')
    ws.write(0, 4, 'first_name')
    ws.write(0, 5, 'last_name')
    ws.write(0, 6, 'title')
    ws.write(0, 7, 'company')
    ws.write(0, 8, 'city')
    ws.write(0, 9, 'state')
    ws.write(0, 10, 'location')
    ws.write(0, 11, 'education')
    ws.write(0, 12, 'num_followers')
    ws.write(0, 13, 'other_info')
    ws.write(0, 14, 'linkedin_url')

    ws.col(0).width = 2000
    ws.col(1).width = 6000
    ws.col(2).width = 6000
    ws.col(3).width = 2000
    ws.col(4).width = 6000
    ws.col(5).width = 6000
    ws.col(6).width = 6000
    ws.col(7).width = 6000
    ws.col(8).width = 6000
    ws.col(9).width = 6000
    ws.col(10).width = 6000
    ws.col(11).width = 10000
    ws.col(12).width = 6000
    ws.col(13).width = 6000
    ws.col(14).width = 12000

    #Experience
    ws = wb.add_sheet('experience')

    ws.write(0, 0, 'our_id')
    ws.write(0, 1, 'linkedin_id')
    ws.write(0, 2, 'first_name')
    ws.write(0, 3, 'last_name')
    ws.write(0, 4, 'exp_company')
    ws.write(0, 5, 'exp_title')
    ws.write(0, 6, 'exp_beg_year')
    ws.write(0, 7, 'exp_beg_month')
    ws.write(0, 8, 'exp_end_year')
    ws.write(0, 9, 'exp_end_month')
    ws.write(0, 10, 'exp_city')
    ws.write(0, 11, 'exp_state')
    ws.write(0, 12, 'exp_country')

    ws.col(0).width = 2000
    ws.col(1).width = 2000
    ws.col(2).width = 6000
    ws.col(3).width = 6000
    ws.col(4).width = 8000
    ws.col(5).width = 8000
    ws.col(6).width = 6000
    ws.col(7).width = 6000
    ws.col(8).width = 6000
    ws.col(9).width = 6000
    ws.col(10).width = 6000
    ws.col(11).width = 6000
    ws.col(12).width = 6000


    #education
    ws = wb.add_sheet('education')

    ws.write(0, 0, 'our_id')
    ws.write(0, 1, 'linkedin_id')
    ws.write(0, 2, 'first_name')
    ws.write(0, 3, 'last_name')
    ws.write(0, 4, 'edu')
    ws.write(0, 5, 'edu_degree')
    ws.write(0, 6, 'edu_beg_year')
    ws.write(0, 7, 'edu_beg_month')
    ws.write(0, 8, 'edu_end_year')
    ws.write(0, 9, 'edu_end_month')
    ws.write(0, 10, 'major')
    ws.write(0, 11, 'city')
    ws.write(0, 12, 'state')
    ws.write(0, 13, 'country')
    ws.write(0, 14, 'activities')

    ws.col(0).width = 2000
    ws.col(1).width = 2000
    ws.col(2).width = 6000
    ws.col(3).width = 6000
    ws.col(4).width = 6000
    ws.col(5).width = 6000
    ws.col(6).width = 6000
    ws.col(7).width = 6000
    ws.col(8).width = 6000
    ws.col(9).width = 6000
    ws.col(10).width = 6000
    ws.col(11).width = 6000
    ws.col(12).width = 6000


    #skill
    ws = wb.add_sheet('skills')

    ws.write(0, 0, 'our_id')
    ws.write(0, 1, 'id')
    ws.write(0, 2, 'first_name')
    ws.write(0, 3, 'last_name')
    ws.write(0, 4, 'skills')
    ws.write(0, 5, 'endo_first_name')
    ws.write(0, 6, 'endo_last_name')
    ws.write(0, 7, 'endo_id')
    ws.write(0, 8, 'endo_title')
    ws.write(0, 9, 'endo_company')
    ws.write(0, 10, 'endo_other_info')

    ws.col(0).width = 2000
    ws.col(1).width = 2000
    ws.col(2).width = 6000
    ws.col(3).width = 6000
    ws.col(4).width = 6000
    ws.col(5).width = 6000
    ws.col(6).width = 6000
    ws.col(7).width = 6000
    ws.col(8).width = 6000
    ws.col(9).width = 6000
    ws.col(10).width = 6000


    #following influencers
    ws = wb.add_sheet('following influencers')

    ws.write(0, 0, 'our_id')
    ws.write(0, 1, 'id')
    ws.write(0, 2, 'first_name')
    ws.write(0, 3, 'last_name')
    ws.write(0, 4, 'inf_first_name')
    ws.write(0, 5, 'inf_last_name')
    ws.write(0, 6, 'inf_title')
    ws.write(0, 7, 'inf_company')
    ws.write(0, 8, 'inf_edu')
    ws.write(0, 9, 'inf_id')
    ws.write(0, 10, 'inf_url')
    ws.write(0, 11, 'inf_num_followers')
    ws.write(0, 12, 'inf_other')

    ws.col(0).width = 2000
    ws.col(1).width = 2000
    ws.col(2).width = 6000
    ws.col(3).width = 6000
    ws.col(4).width = 6000
    ws.col(5).width = 6000
    ws.col(6).width = 12000
    ws.col(7).width = 6000
    ws.col(8).width = 6000
    ws.col(9).width = 6000
    ws.col(10).width = 12000
    ws.col(11).width = 6000
    ws.col(12).width = 6000


    #following school
    ws = wb.add_sheet('following school etc')

    ws.write(0, 0, 'our_id')
    ws.write(0, 1, 'linkedin_id')
    ws.write(0, 2, 'first_name')
    ws.write(0, 3, 'last_name')
    ws.write(0, 4, 'foll_company')
    ws.write(0, 5, 'foll_school')
    ws.write(0, 6, 'foll_group')

    ws.col(0).width = 2000
    ws.col(1).width = 2000
    ws.col(2).width = 6000
    ws.col(3).width = 6000
    ws.col(4).width = 6000
    ws.col(5).width = 6000
    ws.col(6).width = 6000


    #accomplishments
    ws = wb.add_sheet('accomplishments')

    ws.write(0, 0, 'our_id')
    ws.write(0, 1, 'linkedin_id')
    ws.write(0, 2, 'first_name')
    ws.write(0, 3, 'last_name')
    ws.write(0, 4, 'certificate')
    ws.write(0, 5, 'certificate_year')
    ws.write(0, 6, 'certificate_institution')
    ws.write(0, 7, 'award')
    ws.write(0, 8, 'award_year')
    ws.write(0, 9, 'award_institution')
    ws.write(0, 10, 'language')


    ws.col(0).width = 2000
    ws.col(1).width = 2000
    ws.col(2).width = 6000
    ws.col(3).width = 6000
    ws.col(4).width = 6000
    ws.col(5).width = 6000
    ws.col(6).width = 6000
    ws.col(7).width = 6000
    ws.col(8).width = 6000
    ws.col(9).width = 6000
    ws.col(10).width = 6000


    #also viewed
    ws = wb.add_sheet('also viewed')

    ws.write(0, 0, 'our_id')
    ws.write(0, 1, 'id')
    ws.write(0, 2, 'first_name')
    ws.write(0, 3, 'last_name')
    ws.write(0, 4, 'viewed_first_name')
    ws.write(0, 5, 'viewed_last_name')
    ws.write(0, 6, 'viewed_id')
    ws.write(0, 7, 'viewed_title')
    ws.write(0, 8, 'viewed_company')
    ws.write(0, 9, 'viewed_other_info')

    ws.col(0).width = 2000
    ws.col(1).width = 2000
    ws.col(2).width = 6000
    ws.col(3).width = 6000
    ws.col(4).width = 6000
    ws.col(5).width = 6000
    ws.col(6).width = 6000
    ws.col(7).width = 12000
    ws.col(8).width = 6000
    ws.col(9).width = 6000

    wb.save(path)



def _getOutCell(outSheet, colIndex, rowIndex):
    """ HACK: Extract the internal xlwt cell representation. """
    row = outSheet._Worksheet__rows.get(rowIndex)
    if not row: return None

    cell = row._Row__cells.get(colIndex)
    return cell

def setOutCell(outSheet, row, col, value):
    """ Change cell value without changing formatting. """
    # HACK to retain cell style.
    previousCell = _getOutCell(outSheet, col, row)
    # END HACK, PART I

    outSheet.write(row, col, value)

    # HACK, PART II
    if previousCell:
        newCell = _getOutCell(outSheet, col, row)
        if newCell:
            newCell.xf_idx = previousCell.xf_idx
    # END HACK



def edit_file(path, sheetNum, elemList):
    rb = open_workbook(path, formatting_info=True)
    this_sheet = rb.sheet_by_index(sheetNum)
    # read a row
    rowNum = this_sheet.nrows
    #print(this_sheet.nrows)

    wb = copy(rb)

    s = wb.get_sheet(sheetNum)
    i = 0
    row = rowNum


    for elem in elemList:
        setOutCell(s, row, i, elem)
        i+=1
    wb.save(path)
