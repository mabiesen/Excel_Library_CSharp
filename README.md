# Excel_Library_CSharp

Library to pull data from excel workbook and store in an intuitive structure.

#### Basic Use

* The Excel service is instantiated.  A path is passed to open an excel file.
* The service will create a Workbook class object, which contains Worksheet class objects, which contains tables
* Note: Only one table will be automatically created per worksheet, and that is the worksheet's used range

* Once this has occurred, you can access data in the following fashion: 

```
	ExcelService.ExcelWorkbook.worsheetList[worksheet index 0 based].worksheetTables[table list 0 based].dataTable
```

