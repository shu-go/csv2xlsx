csv2xlsx -o test1.xlsx --no-guess test.csv

csv2xlsx -o test2.xlsx test.csv

csv2xlsx -o test3.xlsx --cols text:text,number:number,date:date,time:time,datetime:"datetime(06/01/02 15:04:05)" test.csv

rem csv2xlsx -o test4.xlsx --header 3 --cols text:text,number:number,date:date,time:time,datetime:"datetime(06/01/02 15:04:05)" test.csv
csv2xlsx -o test4.xlsx --header 3 --cols d:text,e:number,f:date,g:time,h:"datetime(06/01/02 15:04:05)" test.csv

csv2xlsx -o test5.xlsx --cols text:text,dummy!number:number,date:date,time:time,datetime:"datetime(06/01/02 15:04:05)" test.csv

rem test1.csv!number is number
rem test2.csv!number is text
rem numbers have red color
csv2xlsx -o test6.xlsx --cols text:text,number:number,test2.csv!number:text,date:date,time:time,datetime:"datetime(06/01/02 15:04:05)"  --number-xlsx [red]#,##0.00 test.csv test2.csv
