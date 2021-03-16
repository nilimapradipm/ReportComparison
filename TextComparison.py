import sys, os, time, difflib
import difflib
import hashlib, sys
import time
import datetime

PATH1 = "./txt/OldVersion/"
PATH2 = "./txt/NewVersion/"
PATH3 = "./txt/CompareResults/"
folder1 = os.listdir(PATH1) # folder containing your files
folder2 = os.listdir(PATH2) # folder containing your files
folder3 = os.listdir(PATH3) # folder containing your files


print("##################### Execution Started #######################")
start_time = time.time()


HtmlHeaderStartingString = '<html>' + '\n' + '<head>' + '\n' + '<h1 style = "background-color:powderblue; text-align:center;font-size:30px;">' + '\n' + '<img src = "maveric.png" align = "right" alt = "Italian Trulli" width = "100" height = "35" >TEXT FILE COMPARISON DASHBOARD</h1>' + '\n' + '<link rel="stylesheet" type="text/css" href="mystyle.css">' + '\n' + '</head>' + '\n' + '<div class="floatLeft" >' + '\n'
HtmlEndString = '</div>'
TableStartString = '\n' + '<table align="center"  CELLSPACING=0 CELLPADDING=5 border="1">' + '</br>'
TableColumnHeaderString = '\n' +  '<tr>' + '<th width="30%">' + 'OLDVERSION_VS_NEWVERSION OVERALL_DIFF_REPORT' + '</th>'  + '<th>' + 'OLDVERSION_VS_NEWVERSION CONTEXT_DIFF_REPORT' + '</th>' + '<th>' +  'OLDVERSION_VS_NEWVERSION FILESIZE' + '</th>' +  '<th>' +  'EXECUTION TIME' + '</th>'  + '<th>' +  'STATUS' + '</th>'  + '</tr>' + '<br/>'


with open("OverallReport.html", 'w') as _file:

    _file.write(HtmlHeaderStartingString)
    _file.write(HtmlEndString)
    _file.write(TableStartString)
    _file.write(TableColumnHeaderString)
for  item1 in folder1:
    for  item2 in folder2:
         if(item1==item2):


           start_time1 = time.time()

           f1 = open(PATH1+item1, 'r').readlines()
           f2 = open(PATH2+item2, 'r').readlines()

           file_stats = os.stat(PATH1+item1)

           diff = difflib.HtmlDiff().make_file(f1, f2, PATH1+item1,
                                       PATH2+item2)


           outputfile = PATH3 + os.path.splitext(item1)[0] + "_overall_diff" + ".html"

           difference_report = open(outputfile, 'w')
           difference_report.write(diff)
           difference_report.close()

           contextdiff = difflib.HtmlDiff().make_file(f1, f2, PATH1 + item1,
                                               PATH2+item2,context=True,numlines=0)
           outputfile1 = PATH3 + os.path.splitext(item1)[0] + "_context_diff" + ".html"

           difference_report = open(outputfile1, 'w')
           difference_report.write(str(contextdiff))
           difference_report.close()

           outputfile1 = PATH1 + os.path.splitext(item1)[0] + ".txt"
           outputfile2 = PATH2 + os.path.splitext(item1)[0] + ".txt"
           files = [outputfile1, outputfile2]


           def md5(fname):
               md5hash = hashlib.md5()
               with open(fname) as handle:  # opening the file one line at a time for memory considerations
                   for line in handle:
                       md5hash.update(line.encode('utf-8'))
               return (md5hash.hexdigest())


           print('Comparing Files:', files[0], 'and', files[1])

           if md5(files[0]) == md5(files[1]):
               print('Matched')
               status = 'Matched'
           else:
               print('Differences')
               status = 'Differences'

           size1 = os.path.getsize(outputfile1)
           size2 = os.path.getsize(outputfile2)
           #print(outputfile1)
           #(size1/1000)
           #print(size2/1000)
           end_time1 = time.time()
           TimeTaken = (end_time1 - start_time1)
           TimeTaken = (end_time1 - start_time1)

           DataString = '\n' + '<tr>' + '<td>' + '<a' + ' href = ' + "http://localhost:63342/pythonProject/txt/CompareResults/" + os.path.splitext(item1)[0] + "_overall_diff" +  ".html" + '>' + os.path.splitext(item1)[
                            0] + "_overall_diff" + '</a>' + '</td>' + '<td>' + '<a' + ' href = ' + "http://localhost:63342/pythonProject/txt/CompareResults/" +  os.path.splitext(item1)[0] + "_context_diff" + ".html" + '>' + os.path.splitext(item1)[
                            0] + "_context_diff" + '</a>' + '</td>' + '<td>' + str(size1 / 1000) + ' vs ' + str(
               size2 / 1000) + ' KB' + '</td>' + '<td>' + str(
               round(TimeTaken, 4)) + ' sec' + '</td>' + '<td>' + status + '</td>' + '</tr>'

           with open("OverallReport.html", 'a') as _file:
                            _file.write(DataString)


end_time = time.time()
TimeTaken=(end_time - start_time)
print('Time Taken For Execution:'+ str(round(TimeTaken,4)))
print("################ Execution Completed in "+str(TimeTaken)+" ###############")

with open("OverallReport.html", 'a') as _file:
 TableEndString = '\n' +  '</table>'
 TableEndHtmlString = '\n' + '</html>'
 OverallExecutionTime = '\n'  + '<p align="center">' 'Overall Execution Time : ' +  str(round(TimeTaken,4)) + ' sec' '</p>' +'\n'
 TimeStampString = '\n' + '<p align="center">' + 'Report Generated TimeStamp : ' + str(datetime.datetime.now()) + '</p>' +'\n' + '<br/>'
 _file.write(TableEndString)
 _file.write(TableEndHtmlString)
 _file.write(OverallExecutionTime)
 _file.write(TimeStampString)
