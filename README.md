# Excel_Library_CSharp

Library to pull data from excel workbook and store in an intuitive structure.


#### Structure

* Excel service contains methods to get excel data and a and Activeworkbook
* Activeworkbook references workbook, which has a name and list of ExcelWorksheets called worsheetList
* Excel Worksheets contain name and a list of ExcelTables
* Excel Table contains one DataTable and an instantiation method for creating data table from 2d array.


#### Basic Use

* The Excel service is instantiated.  A path is passed to open an excel file.
* The service will create a Workbook class object, which contains Worksheet class objects, which contains tables
* Note: Only one table will be automatically created per worksheet, and that is the worksheet's used range

* Once this has occurred, you can access data in the following fashion: 

```
	ExcelService.activeWorkbook.worsheetList[worksheet index 0 based].worksheetTables[table list 0 based].dataTable
```


#### TODO

* As of right now, table only works with strings.  Alter table creation method to identify column type, allow that type.
