import pandas as pd
import sys, os, time
import difflib
import hashlib
import datetime

from pathlib import Path

path_OLD = 'xls/OldVersion/'
path_NEW = 'xls/NewVersion/'
PATH3 = 'xls/CompareResults/'
folder1 = os.listdir(path_OLD) # folder containing your files
folder2 = os.listdir(path_NEW) # folder containing your files
folder3 = os.listdir(PATH3) # folder containing your files

def excel_diff():


    HtmlHeaderStartingString = '<html>' + '\n' + '<head>' + '\n' + '<h1 style = "background-color:powderblue; text-align:center;font-size:30px;">' + '\n' + '<img src = "maveric.png" align = "right" alt = "Italian Trulli" width = "100" height = "35" >EXCEL FILE COMPARISON DASHBOARD</h1>' + '\n' + '<link rel="stylesheet" type="text/css" href="mystyle.css">' + '\n' + '</head>' + '\n' + '<div class="floatLeft" >' + '\n'
    HtmlEndString = '</div>'
    TableStartString = '\n' + '<table align="center"  CELLSPACING=0 CELLPADDING=5 border="1">' + '</br>'
    TableColumnHeaderString = '\n' + '<tr>' + '<th width="30%">' + 'OLDVERSION_VS_NEWVERSION_DIFF_REPORT' + '</th>' + '<th>' + 'OLDVERSION_VS_NEWVERSION FILESIZE' + '</th>' + '<th>' + 'EXECUTION TIME' + '</th>' + '<th>' + 'STATUS' + '</th>' + '</tr>' + '</br>'

    with open("OverallReportExcel.html", 'w') as _file:

        _file.write(HtmlHeaderStartingString)
        _file.write(HtmlEndString)
        _file.write(TableStartString)
        _file.write(TableColumnHeaderString)

    for item1 in folder1:
        for item2 in folder2:
            if (item1 == item2):
              start_time1 = time.time()
              path_OLD1 = path_OLD + item1
              df_OLD = pd.read_excel(path_OLD1).fillna(0)
              #print(df_OLD)
              path_NEW1 = path_NEW + item2
              df_NEW = pd.read_excel(path_NEW1).fillna(0)

    # Perform Diff
    dfDiff = df_OLD.copy()
    Flag = False
    for row in range(dfDiff.shape[0]):
        #print ( dfDiff.shape[0])
        for col in range(dfDiff.shape[1]):
           # print( dfDiff.shape[1])
            value_OLD = df_OLD.iloc[row,col]
            #print(df_OLD.iloc[row,col])
            try:
                value_NEW = df_NEW.iloc[row,col]
             #   print('new' + df_NEW.iloc[row, col])
                if value_OLD==value_NEW:
                    dfDiff.iloc[row,col] = df_NEW.iloc[row,col]
                   # Flag = False
                else:
                    dfDiff.iloc[row,col] = ('{}-->{}').format(value_OLD,value_NEW)
                    Flag = True

            except:
                dfDiff.iloc[row,col] = ('{}-->{}').format(value_OLD, 'NaN')
    if (Flag):
        status = 'Differences'
        print('Differences')
    else:
        status = 'Matched'
        print('Matched')
    # Save output and format
    fname = '{}vs{}.xls'.format(os.path.splitext(item1)[0],os.path.splitext(item2)[0])
    print(fname)

    outputfile = PATH3 + fname
    writer = pd.ExcelWriter(outputfile, engine='xlsxwriter')
    print (path_OLD+item1)
    dfDiff.to_excel(writer, sheet_name='DIFF', index=False)
   # df_NEW.to_excel(writer, sheet_name=os.path.splitext(item1)[0], index=False)
  #  df_OLD.to_excel(writer, sheet_name=os.path.splitext(item2)[0], index=False)

    # get xlsxwriter objects
    workbook  = writer.book
    worksheet = writer.sheets['DIFF']
    worksheet.hide_gridlines(2)

    # define formats
    date_fmt = workbook.add_format({'align': 'center', 'num_format': 'yyyy-mm-dd'})
    center_fmt = workbook.add_format({'align': 'center'})
    number_fmt = workbook.add_format({'align': 'center', 'num_format': '#,##0.00'})
    cur_fmt = workbook.add_format({'align': 'center', 'num_format': '$#,##0.00'})
    perc_fmt = workbook.add_format({'align': 'center', 'num_format': '0%'})
    grey_fmt = workbook.add_format({'font_color': '#E0E0E0'})
    #highlight_fmt = workbook.add_format({'font_color': '#FF0000', 'bg_color':'#B1B3B3'})
    highlight_fmt = workbook.add_format({'font_color': '#FF0000'})
    # set column width and format over columns
    # worksheet.set_column('J:AX', 5, number_fmt)

    # set format over range
    ## highlight changed cells
    worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
                                            'criteria': 'containing',
                                            'value':'-->',
                                            'format': highlight_fmt})
    ## highlight unchanged cells
    worksheet.conditional_format('A1:ZZ1000', {'type': 'text',
                                            'criteria': 'not containing',
                                            'value':'-->',
                                            'format': grey_fmt})

    # save
    writer.save()
    print('Done.')

    # def md5(fname):
    #     md5hash = hashlib.md5()
    #     with open(fname) as handle:  # opening the file one line at a time for memory considerations
    #         for line in handle:
    #             md5hash.update(line.encode('utf-8'))
    #     return (md5hash.hexdigest())
    #
    # print('Comparing Files:', files[0], 'and', files[1])
    #
    # if md5(files[0]) == md5(files[1]):
    #     print('Matched')
    #     status = 'Matched'
    # else:
    #     print('Differences')
    #     status = 'Differences'

    size1 = os.path.getsize(path_OLD1)
    size2 = os.path.getsize(path_NEW1)

    end_time1 = time.time()
    TimeTaken = (end_time1 - start_time1)
    TimeTaken = (end_time1 - start_time1)

    # DataString = '\n' + '<tr>' + '<td>' + '<a' + ' href = ' + "http://localhost:63342/ReportComparison/xls/CompareResults/" + \
    #              os.path.splitext(item1)[0]  + "vs" +  os.path.splitext(item2)[
    #                   0]  + ".xls" + '>' + os.path.splitext(item1)[0]  + "_Diff" +  os.path.splitext(item2)[0]+".xls"'</a>'  '</td>'+ '<td>' + str(size1/1000) + ' vs ' + str(
    #      size2/1000) + ' KB' + '</td>' + '<td>' + str(
    #      round(TimeTaken, 4)) + ' sec' + '</td>' + '<td>' + status + '</td>' + '</tr>'
    DataString = '\n' + '<tr>' + '<td>' + '<a' + ' href = ' + "http://localhost:63342/ReportComparison/xls/CompareResults/" + \
                      os.path.splitext(item1)[0]  + "vs" +  os.path.splitext(item2)[
                      0]  + ".xls" + '>' + os.path.splitext(item1)[0]  + "_Diff" +".xls"'</a>'  '</td>'+ '<td>' + str(size1/1000) + ' vs ' + str(
        size2/1000) + ' KB' + '</td>' + '<td>' + str(
         round(TimeTaken, 4)) + ' sec' + '</td>' + '<td>' + status + '</td>' + '</tr>' + '</br>'
    #print(DataString)
    with open("OverallReportExcel.html", 'a') as _file:
         #_file.flush()
         _file.write(DataString)
        # print(DataString)


def main():
    path_OLD = Path('xls/OldVersion/')
    path_NEW = Path('xls/NewVersion/')
  #  path_OLD = Path('xls/OldVersion/AccountStatement1.xls')
  #  path_NEW = Path('xls/NewVersion/AccountStatement1.xls')
    print("##################### Execution Started #######################")
    start_time = time.time()
    #
    # HtmlHeaderStartingString = '<html>' + '\n' + '<head>' + '\n' + '<h1 style = "background-color:powderblue; text-align:center;font-size:30px;">' + '\n' + '<img src = "maveric.png" align = "right" alt = "Italian Trulli" width = "100" height = "35" >EXCEL FILE COMPARISON DASHBOARD</h1>' + '\n' + '<link rel="stylesheet" type="text/css" href="mystyle.css">' + '\n' + '</head>' + '\n' + '<div class="floatLeft" >' + '\n'
    # HtmlEndString = '</div>'
    # TableStartString = '\n' + '<table align="center"  CELLSPACING=0 CELLPADDING=5 border="1">' + '</br>'
    # TableColumnHeaderString = '\n' + '<tr>' + '<th width="30%">' + 'OLDVERSION_VS_NEWVERSION OVERALL_DIFF_REPORT' + '</th>' + '<th>' + 'OLDVERSION_VS_NEWVERSION FILESIZE' + '</th>' + '<th>' + 'EXECUTION TIME' + '</th>' + '<th>' + 'STATUS' + '</th>' + '</tr>'
    #
    # with open("OverallReport1.html", 'w') as _file:
    #
    #     _file.write(HtmlHeaderStartingString)
    #     _file.write(HtmlEndString)
    #     _file.write(TableStartString)
    #     _file.write(TableColumnHeaderString)

    excel_diff()

    end_time = time.time()
    TimeTaken = (end_time - start_time)
    print('Time Taken For Execution:' + str(round(TimeTaken, 4)))
    print("################ Execution Completed in " + str(TimeTaken) + " ###############")

    with open("OverallReportExcel.html", 'a') as _file:
            TableEndString = '\n' + '</table>'
            TableEndHtmlString = '\n' + '</html>'
            OverallExecutionTime = '\n' + '<p align="center">' 'Overall Execution Time : ' + str(
                round(TimeTaken, 4)) + ' sec' '</p>' + '\n'
            TimeStampString = '\n' + '<p align="center">' + 'Report Generated TimeStamp : ' + str(
                datetime.datetime.now()) + '</p>' + '\n' + '<br/>'
            _file.write(TableEndString)
            _file.write(TableEndHtmlString)
            _file.write(OverallExecutionTime)
            _file.write(TimeStampString)


if __name__ == '__main__':
    main()