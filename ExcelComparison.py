import numpy as np
import pandas as pd
import sys, os, time
import difflib
import hashlib
import datetime

from pathlib import Path

#from TextComparison import start_time

path_ExcelOLD = 'xls/OldVersion/'
path_ExcelNEW = 'xls/NewVersion/'
PATH3 = 'xls/CompareResults/'
folder1 = os.listdir(path_ExcelOLD) # folder containing your files
folder2 = os.listdir(path_ExcelNEW) # folder containing your files
folder3 = os.listdir(PATH3) # folder containing your files


def excel_diff():
    print("##################### Execution Started #######################")
    start_time = time.time()

    HtmlHeaderStartingString = '<html>' + '\n' + '<head>' + '\n' + '<h1 style = "background-color:powderblue; text-align:center;font-size:30px;">' + '\n' + '<img src = "maveric.png" align = "right" alt = "Italian Trulli" width = "100" height = "35" >EXCEL FILE COMPARISON DASHBOARD</h1>' + '\n' + '<link rel="stylesheet" type="text/css" href="mystyle.css">' + '\n' + '</head>' + '\n' + '<div class="floatLeft" >' + '\n'
    HtmlEndString = '</div>'
    TableStartString = '\n' + '<table align="center"  CELLSPACING=0 CELLPADDING=5 border="1">' + '</br>'
    TableColumnHeaderString = '\n' + '<tr>' + '<th width="30%">' + 'OLDVERSION_VS_NEWVERSION_DIFF_REPORT' + '</th>' + '<th>' + 'OLDVERSION_VS_NEWVERSION FILESIZE' + '</th>' + '<th>' + 'EXECUTION TIME' + '</th>' + '<th>' + 'STATUS' + '</th>' + '<th>' + 'TOTAL DIFFERENCES' + '</th>' + '</tr>' + '</br>'

    with open("OverallReportExcel.html", 'w') as _file:

        _file.write(HtmlHeaderStartingString)
        _file.write(HtmlEndString)
        _file.write(TableStartString)
        _file.write(TableColumnHeaderString)

    for item1 in folder1:
        for item2 in folder2:
            if (item1 == item2):
                start_time1 = time.time()
                path_OLD1 = path_ExcelOLD + item1
                # df_OLD = pd.read_excel(path_OLD1).fillna(0)
                df1 = pd.read_excel(path_OLD1, header=None)
                #df1.to_excel("xls/CompareResults/b.xls", index=True)
                # print(df_OLD)
                path_NEW1 = path_ExcelNEW + item2
                # df_NEW = pd.read_excel(path_NEW1).fillna(0)
                df2 = pd.read_excel(path_NEW1, header=None)

                df3 = df1.append(pd.Series(['-----------------------------']), ignore_index=True)
               # print(df1)

                result = df3.append(df2)

                result = result.drop_duplicates(keep=False)
               # result.loc[~result.index.isin(df2.index), 'status']
               #result.loc[~result.index.isin(df1.index), 'status']
               # idx = df3.stack().groupby(level=[0, 1]).nunique()
                #result.loc[idx.mask(idx <= 1).dropna().index.get_level_values(0), 'status'] = 'MODIFIED'
                #print(result)
                total:int = len(result) - 1
                number:int = 0
                #print(Total Differences)
                #print(number)

                if (total == 0):
                     status = 'Matched'
                else:
                    status = 'Differences'
                print(status)


                fname = '{}vs{}.xls'.format(os.path.splitext(item1)[0], os.path.splitext(item2)[0])

                writer = pd.ExcelWriter("xls/CompareResults/" + fname, engine='xlsxwriter')
                result.to_excel(writer, sheet_name='DIFF', index=True)
                df1.to_excel(writer, sheet_name='OldVersion', index=True)
                df2.to_excel(writer, sheet_name='NewVersion', index=True)

                # get xlsxwriter objects
                workbook = writer.book

                worksheet = writer.sheets['DIFF']
                worksheet = writer.sheets['OldVersion']
                worksheet = writer.sheets['NewVersion']
                worksheet.hide_gridlines(2)
                worksheet.set_default_row(15)

                writer.save()

                print('\nDone.\n')


                size1 = os.path.getsize(path_OLD1)
                size2 = os.path.getsize(path_NEW1)

                end_time1 = time.time()

                TimeTaken = (end_time1 - start_time1)



                DataString = '\n' + '<tr>' + '<td>' + '<a' + ' href = ' + "http://localhost:63342/ReportComparison/xls/CompareResults/" + \
                 os.path.splitext(item1)[0] + "vs" + os.path.splitext(item2)[
                     0] + ".xls" + '>' + os.path.splitext(item1)[0] + "_Diff" + ".xls"'</a>'  '</td>' + '<td>' + str(
        size1 / 1000) + ' vs ' + str(
        size2 / 1000) + ' KB' + '</td>' + '<td>' + str(
        round(TimeTaken, 4)) + ' sec' + '</td>' + '<td>' + status + '</td>' + '<td>' + str(total) + '</td>''</tr>' + '</br>'
                #print(DataString)
                with open("OverallReportExcel.html", 'a') as _file:

                   _file.write(DataString)
    # print(DataString)
            end_time = time.time()
            TimeTaken = (end_time - start_time)
            #print('Time Taken For Execution:' + str(round(TimeTaken, 4)))
            #print("################ Execution Completed in " + str(TimeTaken) + " ###############")


def main():
    #path_OLD = Path('xls/OldVersion/AccountStatement1.xls')
    #path_NEW = Path('xls/NewVersion/AccountStatement1.xls')
    #print("##################### Execution Started #######################")
    start_time = time.time()
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