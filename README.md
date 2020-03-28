# Convert-CsvToXlsx
This is a PowerShell function that converts a given text file to an Excel Spreadsheet (*.xlsx).

The function reads the contents of a given file and then outputs the contents in an xlsx format.

It automatically creates thin borders for the headers and fills the cells the headers are in with a light gray color.

## Parameters
### InputFilePath
[REQUIRED]

A file that contains the data to be converted to Excel xlsx format.

### Delimiter
[OPTIONAL]

The way that the columns are delimited in the InputFile.

### Header
[OPTIONAL]

The header(s) for the Excel Worksheet. If no header(s) is/are specified then the first line in the InputFile is used as the header(s). Multiple values can be used here and need to be separated by "," or ";".

# Usage
## Example 1
```powershell
Convert-CsvToXlsx -InputFilePath C:\temp\ServersByOs.csv -Delimiter ","
```
This will return a file called C:\temp\ServersByOs.xlsx. The first row in C:\temp\ServersByOs.csv will be used as the headers. The columns will be separated by the delimiter ",".

The data would look like this in C:\temp\ServersByOs.csv:

```text
ServerName,OSVersion,InstalledDate
firstserver01,Windows 2012,03/05/2016
secondserver02,RedHat 5,01/02/2015
```

The data would look like this in C:\temp\ServersByOs.xlsx:

```text  
  |   ServerName   |   OSVersion  | InstalledDate |
  -------------------------------------------------
  | firstserver01  | Windows 2012 | 03/05/2016    |
  | secondserver02 | RedHat 5	  | 01/02/2015    |
```

## Example 2
```powershell
Convert-CsvToXlsx -InputFilePath C:\temp\ServersByOs.csv -Delimiter ";" -Header "ServerName,OSVersion,InstalledDate"
```
This will return a file called C:\temp\ServersByOs.xlsx. The headers specified in Header will be used as the headers. The columns will be separated by the delimiter ";". 

The data would look like this in C:\temp\ServersByOs.csv:
```text  
"firstserver01";"Windows";"2012;03/05/2016"
"secondserver02";"RedHat 5";"01/02/2015"
```

The data would look like this in C:\temp\ServersByOs.xlsx:
```text  
  |   ServerName   |   OSVersion  | InstalledDate |
  -------------------------------------------------
  | firstserver01  | Windows 2012 | 03/05/2016    |
  | secondserver02 | RedHat 5	  | 01/02/2015    |
```

##Example 3
```powershell
Convert-CsvToXlsx -InputFilePath C:\temp\Servers.csv
```
This will return a file called C:\temp\Servers.xlsx.
The first row in C:\temp\Servers.csv will be used as the headers.
The will be only one column in the Excel Worksheet because there is no Delimiter specified.

The data would look like this in C:\temp\Servers.csv:

```text
ServerName
firstserver01
secondserver02,secondserver02.fqdn.com
```

The data would look like this in C:\temp\Servers.xlsx:

```text  
  |   			ServerName	   |
  -----------------------------------------
  | firstserver01  			   |
  | secondserver02,secondserver02.fqdn.com |
```

# Version
1.1

# License
[GNU General Public License V3](https://www.gnu.org/licenses/gpl-3.0.en.html)
