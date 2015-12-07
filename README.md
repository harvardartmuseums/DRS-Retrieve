# DRS Retrieve

This is a simple Windows console application for fetching image files from the Harvard DRS.


## Usage

```
DRSRetrieve.exe filename [/S:server] [/M:module] [/F:filesize] [/N:filenameformat] [/A]
```

Command line arguments

```
filename: Specifies the input text file containing items to retrieve
/S: Specifies the ODBC connection to the database server
/M: Specifies the module to search
/F: Specifies the filesize to retrieve
/N: Specifies the format of the filename
/A: Retrieves all files for each item in the input file
```

## Example

DRSRetrieve MyObjectList.txt /s:Museum /m:Objects /f:full /n:2