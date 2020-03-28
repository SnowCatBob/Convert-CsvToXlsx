.SYNOPSIS
  Converts a given text file to an Excel Spreadsheet (*.xlsx)
  
.DESCRIPTION
  The function reads the contents of a given file and then outputs the contents in an xlsx format.
  It automatically creats thin borders for the headers and fills the cells the headers are in with a light gray color.
  
.PARAMETER InputFilePath
    [REQUIRED]
    A file that contains the data to be converted to Excel xlsx format.
	
.PARAMETER Delimiter
    [OPTIONAL]
    The way that the columns are delimited in the InputFile.
	For example if you have a .csv, the delimiter would most likely be "," or ";" but the delimiter can be anything you have used in the InputFile to separate the collumns.
	
.PARAMETER Header
    [OPTIONAL]
    The header(s) for the Excel Worksheet.
	If no header(s) is/are specified then the first line in the InputFile is used as the header(s).
	Multiple values can be used here and need to be separated by "," or ";".

.INPUTS
  A file containing the data to be converted to an Excel Spreadsheet (*.xlsx).
  
.OUTPUTS
  An Excel .xlsx file with the same name as the InputFile will be created in the same path as the InputFile.
  
.NOTES
  Version:          1.1
  Author:           SnowCatBob (SnowCatBob01@gmail.com)
  Project Location: https://github.com/SnowCatBob/Convert-CsvToXlsx
  
.EXAMPLE
  Convert-CsvToXlsx -InputFilePath C:\temp\ServersByOs.csv -Delimiter ","
  This will return a file called C:\temp\ServersByOs.xlsx.
  The first row in C:\temp\ServersByOs.csv will be used as the headers.
  The columns will be separated by the delimiter ",".
  The data would look like this in C:\temp\ServersByOs.csv:
  
  ServerName,OSVersion,InstalledDate
  firstserver01,Windows 2012,03/05/2016
  secondserver02,RedHat 5,01/02/2015
  
  The data would look like this in C:\temp\ServersByOs.xlsx:
  
  |   ServerName   |   OSVersion  | InstalledDate |
  -------------------------------------------------
  | firstserver01  | Windows 2012 | 03/05/2016    |
  | secondserver02 | RedHat 5	  | 01/02/2015    |
  
.EXAMPLE
  Convert-CsvToXlsx -InputFilePath C:\temp\ServersByOs.csv -Delimiter ";" -Header "ServerName,OSVersion,InstalledDate"
  This will return a file called C:\temp\ServersByOs.xlsx.
  The headers specified in Header will be used as the headers.
  The columns will be separated by the delimiter ";".
  The data would look like this in C:\temp\ServersByOs.csv:
  
  firstserver01;Windows 2012;03/05/2016
  secondserver02;RedHat 5;01/02/2015
  
  The data would look like this in C:\temp\ServersByOs.xlsx:
  
  |   ServerName   |   OSVersion  | InstalledDate |
  -------------------------------------------------
  | firstserver01  | Windows 2012 | 03/05/2016    |
  | secondserver02 | RedHat 5	  | 01/02/2015    |
  
.EXAMPLE
  Convert-CsvToXlsx -InputFilePath C:\temp\Servers.csv
  This will return a file called C:\temp\Servers.xlsx.
  The first row in C:\temp\Servers.csv will be used as the headers.
  The will be only one column in the Excel Worksheet because there is no Delimiter specified.
  The data would look like this in C:\temp\Servers.csv:
  
  ServerName
  firstserver01
  secondserver02,secondserver02.fqdn.com
  
  The data would look like this in C:\temp\Servers.xlsx:
  
  |   			ServerName				   |
  ------------------------------------------
  | firstserver01  						   |
  | secondserver02,secondserver02.fqdn.com |