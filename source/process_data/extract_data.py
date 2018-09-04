#coding=utf-8
import xlrd
import openpyxl
import sys
import math
if sys.getdefaultencoding() != 'utf-8':
    reload(sys)
    sys.setdefaultencoding('utf-8')



##抽取特殊证素
def extract_special_syndrome_element_data(syndrome_element, data_path, colume_number, medicine_start, medicine_end, threshold):
    data = xlrd.open_workbook(data_path)
    table = data.sheets()[0]
    nrows = table.nrows
    workbook = openpyxl.Workbook()
    worksheet = workbook.create_sheet(title=syndrome_element)
    #生成第一行数据
    for i in xrange(1, 13):
        worksheet.cell(1, i, table.cell(0, i-1).value.encode('utf-8'))
    worksheet.cell(1, column_number, table.cell(0, colume_number).value.encode('utf-8'))
    for i in xrange(0, medicine_end-medicine_start+1):
        worksheet.cell(1, 14+i, table.cell(0, medicine_start + i-1).value.encode('utf-8'))

    worksheet_nraws = 2
    #生成其他行的数据
    print table.cell(1, column_number)
    for i in xrange(1, nrows):
        if math.fabs(table.cell(i, column_number).value - 1.0) < 0.00001:
            for j in xrange(1, 13):
                worksheet.cell(worksheet_nraws, j).value = table.cell(i, j-1).value
            worksheet.cell(worksheet_nraws, column_number).value = table.cell(i, colume_number).value
            for j in xrange(0, medicine_end - medicine_start + 1):
                worksheet.cell(worksheet_nraws, 14 + j).value = table.cell(i, medicine_start + j - 1).value
            worksheet_nraws += 1

    use_cols = []
    for ncol in xrange(14, worksheet.max_column + 1):
        temp = 0
        for nrow in xrange(2, worksheet.max_row + 1):
            temp += worksheet.cell(nrow, ncol).value
        if math.fabs(temp-0.0) > threshold:
            use_cols.append(ncol)
    print use_cols
    print use_cols.__len__()

    worksheet_delete_useless_col = workbook.create_sheet(title="delete_useless")
    #前面13行保留
    for i in xrange(1, worksheet.max_row+1):
        for j in xrange(1, 14):
            worksheet_delete_useless_col.cell(i, j).value = worksheet.cell(i, j).value
        for j in xrange(0, use_cols.__len__()):
            worksheet_delete_useless_col.cell(i, 14+j).value = worksheet.cell(i, use_cols[j]).value

    workbook.save("phlegm_with_medicines.xlsx")
    #抽取出所有该证素为1的行，并且保存到新的sheet中

# data = xlrd.open_workbook("../../data/traditional_chinese_medicine_data.xlsx")
    # table = data.sheets()[0]
    # nrows = table.nrows
    # ncols = table.ncols
    # print nrows
    # print ncols
    # cell_A1 = table.cell(0, 178).value
    # cell_C4 = table.cell(0, 52).value
    # print cell_A1
    # print cell_C4
if __name__ == '__main__':
    data_path = "../../data/traditional_chinese_medicine_data.xlsx"
    syndrome_element = "phlegm_with_medicines"
    column_number = 13
    medicine_start = 180
    medicine_end = 554
    threshold = 0.5
    extract_special_syndrome_element_data(syndrome_element, data_path, column_number, medicine_start, medicine_end, threshold)



