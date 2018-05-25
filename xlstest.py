import xlsxwriter
import os.path
import datetime

# check if file exists
date=str(datetime.datetime.now().strftime("%Y/%m/%d"))

file=datetime.datetime.now().strftime("%B")
file+='-'+datetime.datetime.now().strftime("%Y")+'-speedtests.xlsx'

if os.path.isfile(file) == True:
    print("File ",file," exists")
else:
    print("File doesn't exists, creating it")
    workbook=xlsxwriter.Workbook(file)

    #creating worksheets
    lqd_worksheet=workbook.add_worksheet('LQD')
    jtl_worksheet=workbook.add_worksheet('JTL')
    saf_worksheet=workbook.add_worksheet('SAF')

    #morning text format
    cellB1=workbook.add_format()
    cellB1.set_italic()
    cellB1.set_underline()
    cellB1.set_bold()
    cellB1.set_bg_color('yellow')

    #BOLD big title fonts
    bold_bigTitle_format=workbook.add_format()
    bold_bigTitle_format.set_bold()
    bold_bigTitle_format.set_size(14)

    #add morning and afternoon text
    jtl_worksheet.write('B1','Morning',cellB1)
    saf_worksheet.write('B1','Morning',cellB1)
    lqd_worksheet.write('B1','Morning',cellB1)
    jtl_worksheet.write('O1','Afternoon',cellB1)
    lqd_worksheet.write('O1','Afternoon',cellB1)
    saf_worksheet.write('O1','Afternoon',cellB1)

    #create borders and titles
    bold_bigTitle_format.set_align('center')
    jtl_worksheet.merge_range('E2:H2','JTL LINK SPEEDTEST',bold_bigTitle_format)
    jtl_worksheet.merge_range('R2:U2','JTL LINK SPEEDTEST',bold_bigTitle_format)

    #Dark border format
    dark_border_format=workbook.add_format()
    dark_border_format.set_bold()
    dark_border_format.set_size(14)
    dark_border_format.set_border(2)
    dark_border_format.set_align('center')

    #BOLD TITLE WITH NORMAL BORDER FORMAT
    bold_title_normal_border=workbook.add_format()
    bold_title_normal_border.set_bold()
    bold_title_normal_border.set_size(12)
    bold_title_normal_border.set_border(1)
    bold_title_normal_border.set_align('center')

    #morning table
    jtl_worksheet.merge_range('C3:D3','UK',dark_border_format)
    jtl_worksheet.merge_range('E3:F3','US',dark_border_format)
    jtl_worksheet.merge_range('G3:H3','EUROPE',dark_border_format)
    jtl_worksheet.merge_range('I3:J3','NAIROBI',dark_border_format)
    #DATE raw titles
    jtl_worksheet.write('B4','DATE',bold_title_normal_border)
    jtl_worksheet.write('C4','Download',bold_title_normal_border)
    jtl_worksheet.write('D4','Upload',bold_title_normal_border)
    jtl_worksheet.write('E4','Download',bold_title_normal_border)
    jtl_worksheet.write('F4','Upload',bold_title_normal_border)
    jtl_worksheet.write('G4','Download',bold_title_normal_border)
    jtl_worksheet.write('H4','Upload',bold_title_normal_border)
    jtl_worksheet.write('I4','Download',bold_title_normal_border)
    jtl_worksheet.write('J4','Upload',bold_title_normal_border)
    jtl_worksheet.write('K4','Remarks',bold_title_normal_border)
    jtl_worksheet.write('L4','By',bold_title_normal_border)
    
    #evening table
    jtl_worksheet.merge_range('P3:Q3','UK',dark_border_format)
    jtl_worksheet.merge_range('R3:S3','US',dark_border_format)
    jtl_worksheet.merge_range('T3:U3','EUROPE',dark_border_format)
    jtl_worksheet.merge_range('V3:W3','NAIROBI',dark_border_format)
    #DATE raw titles
    jtl_worksheet.write('O4','DATE',bold_title_normal_border)
    jtl_worksheet.write('P4','Download',bold_title_normal_border)
    jtl_worksheet.write('Q4','Upload',bold_title_normal_border)
    jtl_worksheet.write('R4','Download',bold_title_normal_border)
    jtl_worksheet.write('S4','Upload',bold_title_normal_border)
    jtl_worksheet.write('T4','Download',bold_title_normal_border)
    jtl_worksheet.write('U4','Upload',bold_title_normal_border)
    jtl_worksheet.write('V4','Download',bold_title_normal_border)
    jtl_worksheet.write('W4','Upload',bold_title_normal_border)
    jtl_worksheet.write('X4','Remarks',bold_title_normal_border)
    jtl_worksheet.write('Y4','By',bold_title_normal_border)
    
    
    workbook.close()
