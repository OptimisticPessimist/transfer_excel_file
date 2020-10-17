# transfer_excel_file
The file is corrupted when transfer another excel file.

I want post from a excel data to another excel cells.  
BUT this code corrupts the file.  
When the corrupted file was changed file extension from `.xlsx` to `.xls`, I can open it.

`openpyxl` couldn't write `.xls`, right?  
I can transform value of it, but I can't open it. I confirmed it in the test.


## Environments
- Python 3.7
- openpyxl 3.0.5
- Microsoft Office 365 Excel version 2009