csv2xlsx -o test1.xlsx --no-guess test.csv

csv2xlsx -o test2.xlsx test.csv

csv2xlsx -o test3.xlsx --cols text:text,number:number,date:date,time:time,datetime:"datetime(06/01/02 15:04:05)" test.csv
