# DataTableToXlsPdf
```
______      _      _____     _     _    _____    __   ___     ______   _  __ 
|  _  \    | |    |_   _|   | |   | |  |_   _|   \ \ / / |    | ___ \ | |/ _|
| | | |__ _| |_ __ _| | __ _| |__ | | ___| | ___  \ V /| |___ | |_/ /_| | |_ 
| | | / _` | __/ _` | |/ _` | '_ \| |/ _ \ |/ _ \ /   \| / __||  __/ _` |  _|
| |/ / (_| | || (_| | | (_| | |_) | |  __/ | (_) / /^\ \ \__ \| | | (_| | |  
|___/ \__,_|\__\__,_\_/\__,_|_.__/|_|\___\_/\___/\/   \/_|___/\_|  \__,_|_|  
                                                                                                                                                         
```

## OVERVIEW:
A class library for taking a datatable and converting it to a Microsoft Excel XLSX or a PDF file.  The PDF file generated using the Microsoft Excel export, not PrintToPdf!

*** THIS DOES REQUIRE MICROSOSFT EXCEL TO BE INSTALLED ***

USAGE:
Look at the test but basically it's this...
```
DataTableToXlsPdf.ToFile(dataTable, file);
File is a FileInfo.  The extension either .PDF or something else...  Something else it will save as an excel file (.XLSX)
```