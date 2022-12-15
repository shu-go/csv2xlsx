csv2xlsx -o test1.xlsx --no-guess test.csv

csv2xlsx -o test2.xlsx test.csv

csv2xlsx -o test3.xlsx --cols text:text,number:number,date:date,time:time,datetime:"datetime(06/01/02 15:04:05)" test.csv

rem csv2xlsx -o test4.xlsx --header 3 --cols text:text,number:number,date:date,time:time,datetime:"datetime(06/01/02 15:04:05)" test.csv
csv2xlsx -o test4.xlsx --header 3 --cols d:text,e:number,f:date,g:time,h:"datetime(06/01/02 15:04:05)" test.csv
