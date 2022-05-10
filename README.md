# automate-excel-with or without DB

this code has been written for the use of automating important excel files. 
instead of buying a shelf product to manage company licenses we use this simple code.
"convert.py" - will allow to connect to DB(mySQL) and it has automatic backup and it copies the excel data directly and then creating new file with expiring licenses and sending it via mail (gmail at the moment) - put  your own gmail details and it will work.
"lic-expire.py" - just using pandas for getting only expired licenses and sending them via mail, no DB will be included.

to use convert.py there are some needs:
1. DB (mysql)
2. excel file 
3. change header names of excel and DB and it will work.
