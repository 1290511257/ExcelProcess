#!/usr/bin/python
#-*-coding:utf-8 -*-
import xlrd
import sys
import openpyxl
import datetime
import difflib

# print difflib.SequenceMatcher(None,'付款：法律顾问费（报中国监证会当地派出机构辅导）-北京大成（广州）13437050','北京大成（广州）律师事务所').quick_ratio()

# to confirm it is the same day
def confirm_date(str1,str2,date_weigth):
    #print 'begin process date',str1,str2,str1[4:5],str1[4:5] == '0'
    year = str1[0:4]
    mounth = str1[5:6] if str1[4:5] == '0' else str1[4:6] 
    day = str1[7:8] if str1[6:7] == '0' else str1[6:8]
    str1 = year + "/" + mounth + "/" + day
    #print str1,str2,str1==str2,'weigth =',date_weigth if str1 == str2 else 0,str1[6:7]
    return date_weigth if str1 == str2 else 0

# confirm have the same amount
def confirm_amount(str1,str2):
    # print 'amount1 =',str1,'amount2 =',str2
    return str1 == str2

# if the summary has more than 20% same content
def get_summary_equal_rate(str1,str2):
    return difflib.SequenceMatcher(lambda x: x=="有限公司",str1,str2).quick_ratio()

def ignore_empty(str1):
    return "str" if str1 is None else str1

# init
process_excel_file_path = 'D:\Program Files\Desktop\infomation.xlsx'
workbook = openpyxl.load_workbook(process_excel_file_path);

#加权权重
date_weight = 0.1
summary_weight = 0.3
object_weight = 0.6

worksheet1 = workbook.worksheets[0]
worksheet2 = workbook.worksheets[1]
worksheet_result = workbook.worksheets[2]

worksheet1['E2'].value = '匹配行'
worksheet1['F2'].value = '匹配度'
worksheet1['G2'].value =  '建议'
#rows = worksheet1.rows
#columns = worksheet1.columns

row_index_sheet1 = 0

for date_cell in worksheet1['A']:

    if date_cell.value is None:
        continue
    row_index_sheet1 = row_index_sheet1 + 1
    summary = worksheet1['B'+str(row_index_sheet1)].value
    object_1 =  worksheet1['C'+str(row_index_sheet1)].value
    amount = worksheet1['D'+str(row_index_sheet1)].value
    print "begin process ",row_index_sheet1," row data =",date_cell.value,"summary =",summary,"object = ",object_1," amount =",amount
    
    if row_index_sheet1 < 3  :
        continue
    amount_column_sheet2 = 'G' if amount > 0 else 'F'
    row_index_sheet2 = 0
    max_equal_rate = 0.0
    second_max_equal_rate = 0.0

    #sheet2中的匹配行索引
    most_match_row_index_in_worksheet2 = 0
    second_match_row_index_in_worksheet2 = 0

    notify_row = ''
    notify_message = ''

    for date_cell_compare in worksheet2['A']: # search the same data from worksheet2
        row_index_sheet2 = row_index_sheet2 + 1
        if date_cell_compare is None:
            continue
        # if row_index_sheet2 == 6 :
        #     print '=====> amount =',abs(amount),'sheet2 amount =',worksheet2[amount_column_sheet2+str(row_index_sheet2)].value,'result =',abs(amount) == worksheet2[amount_column_sheet2+str(row_index_sheet2)].value
        # print amount_column_sheet2
        #if confirm_date(date_cell.value,date_cell_compare.value) :#check deal date,if true then check amount
        if confirm_amount(abs(amount),worksheet2[amount_column_sheet2+str(row_index_sheet2)].value) :
            summary = ignore_empty(summary)
            object_1 = ignore_empty(object_1)
            temp_equal_rate = confirm_date(str(date_cell.value),str(date_cell_compare.value),date_weight) + get_summary_equal_rate(summary,worksheet2['C'+str(row_index_sheet2)].value)*summary_weight + get_summary_equal_rate(object_1,worksheet2['C'+str(row_index_sheet2)].value)*object_weight  # caculate match rate
            if max_equal_rate <= temp_equal_rate:
                second_max_equal_rate = max_equal_rate
                max_equal_rate = temp_equal_rate
                second_match_row_index_in_worksheet2 = most_match_row_index_in_worksheet2
                most_match_row_index_in_worksheet2 = row_index_sheet2
            else:
                second_max_equal_rate = second_max_equal_rate if second_max_equal_rate > temp_equal_rate else temp_equal_rate
                second_match_row_index_in_worksheet2 = second_match_row_index_in_worksheet2 if second_max_equal_rate > temp_equal_rate else row_index_sheet2
            #print "worksheet1 row",row_index_sheet1,"with worksheet2 row",row_index_sheet2,"was match date and amount,max equal rate update to",max_equal_rate,'second equal rate update to',second_max_equal_rate
            #print '>>>>match rate max =',get_summary_equal_rate(summary,worksheet2['C'+str(row_index_sheet2)].value),'  rate2 =',get_summary_equal_rate(object_1,worksheet2['C'+str(row_index_sheet2)].value)
            #print row_index_sheet1,list(worksheet1['D'])[row_index_sheet1]
    
    
    if most_match_row_index_in_worksheet2 < 3:
        print 'worksheet1 row',row_index_sheet1,' has no match row in worksheet2'
        continue
    
    print 'worksheet1 row',row_index_sheet1,'most match in worksheet2 row',most_match_row_index_in_worksheet2,'second match row',second_match_row_index_in_worksheet2

    if (max_equal_rate - second_max_equal_rate) > 0.25 : #如果与第二组匹配项差值大于0.25，则认定匹配成功
        print 'MATCH SUCCESS. max equal rate =',max_equal_rate
    elif max_equal_rate < 0.3:
        notify_message = '匹配度过低，请check'
        print 'MATCH FAILED for match date too low. max match rate =',max_equal_rate
    else:
        notify_message = '存在近似匹配行' + str(second_match_row_index_in_worksheet2) + '，请check'
        print 'MATCH SUCCESS but exist approximate data. max match rate =',max_equal_rate
        
    # write information 
    #sheet1

    worksheet1['E'+str(row_index_sheet1)].value = most_match_row_index_in_worksheet2
    worksheet1['F'+str(row_index_sheet1)].value = max_equal_rate

    if notify_message != '':
        # print notify_message
        worksheet1['G'+str(row_index_sheet1)].value = notify_message

    if row_index_sheet1 == 20:
        print '=====>',confirm_date(str(worksheet1['A20'].value),str(worksheet2['A20'].value),date_weight),confirm_date(str(worksheet1['A20'].value),str(worksheet2['A21'].value),date_weight),str(worksheet1['A20'].value),str(worksheet2['A20'].value),str(worksheet2['A21'].value)
    #result sheet
    worksheet_result['A'+str(row_index_sheet1)].value =  worksheet2['A'+str(most_match_row_index_in_worksheet2)].value
    worksheet_result['B'+str(row_index_sheet1)].value =  worksheet2['B'+str(most_match_row_index_in_worksheet2)].value
    worksheet_result['C'+str(row_index_sheet1)].value =  worksheet2['C'+str(most_match_row_index_in_worksheet2)].value
    worksheet_result['D'+str(row_index_sheet1)].value =  worksheet2['D'+str(most_match_row_index_in_worksheet2)].value
    worksheet_result['E'+str(row_index_sheet1)].value =  worksheet2['E'+str(most_match_row_index_in_worksheet2)].value
    worksheet_result['F'+str(row_index_sheet1)].value =  worksheet2['F'+str(most_match_row_index_in_worksheet2)].value
    worksheet_result['G'+str(row_index_sheet1)].value =  worksheet2['G'+str(most_match_row_index_in_worksheet2)].value
    worksheet_result['M'+str(row_index_sheet1)].value =  worksheet1['A'+str(row_index_sheet1)].value
    worksheet_result['N'+str(row_index_sheet1)].value =  worksheet1['B'+str(row_index_sheet1)].value
    worksheet_result['O'+str(row_index_sheet1)].value =  worksheet1['C'+str(row_index_sheet1)].value
    worksheet_result['P'+str(row_index_sheet1)].value =  worksheet1['D'+str(row_index_sheet1)].value
    
workbook.save('result.xlsx')
