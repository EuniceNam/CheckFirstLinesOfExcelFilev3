# CheckFirstLinesOfExcelFilev3
1st lines of excel workbook (each sheet) v2

### Motivation
I had to inspect 2 types of spreadsheet files in my part-time job.  
**type 1**: About 7 thousand in number, each has 54 sheets xlsx files.  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  Each sheet is pasted copy of webpage by a person.  
&nbsp;&nbsp;&nbsp;&nbsp;  -> Check if first lines of sheet follow the restriction: Yv0.3.1.py  
**type 2**: About several tens of thousand in number, one-sheet .xls but html files.  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  Each file is downloaded by a person. Plus .xlsx files are made by the person if to-be-downloaded item has special characteristics  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  i.e. address change.  
&nbsp;&nbsp;&nbsp;&nbsp;  -> Check if the file has allowed character: YSimple0.0.12.py  
Both types of files were named by people.  
&nbsp;&nbsp;&nbsp;&nbsp;  -> Check if files are named in right way, and nothing is ommited or duplicatied: LstCpr  
Both data are gradually generated.  
&nbsp;&nbsp;&nbsp;&nbsp;  -> config.json for YVx.x.x.py  
  
### File
To check the content of multiple 54-sheet xlsx files  
**Yv (main)**: Check if the first line of each sheet of each xlsx file follows the restriction  
**Part**: function and fixed restriction(cform)  
**config**: designate file path, range, etc.  
  
**Sheets2CSV**: convert each sheets of an xlsx file to csv  
**LstCpr**: Check if files are named in right way (as written in list.csv), if existing files have wrong content (that should be exchanged)  
