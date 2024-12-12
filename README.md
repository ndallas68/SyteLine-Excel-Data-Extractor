# SyteLine-Excel-Data-Extractor
The functionality of this software is to provide a means of running excel sheets out of Infor at any given time without manual intervention. By setting this utility up with the correct parameters in the task scheduler, data could be pulled from the system and put in any location at any time.

# Argument Documentation for ExcelDataExtractor

### Required
==================

Argument 1: Client URL with ConfigGroup

Argument 2: Username

Arguemnt 3: Password

Argument 4: Configuration

Argument 5: Must be the filename and location... Note that the filename will be appended with the date and .csv.

	Example of argument: \\sty-fs-1\users\public\CustomerOrderLines
 
	Program Output: \\sty-fs-1\users\public\CustomerOrderLines <Date>.csv
 
Argument 6: Name of the IDO

Argument 7: List of Properties as a comma delimited string, or * for all properties in IDO.

	Example of argument: "EmpNum,Name"
 
Argument 8: Record Cap

	Must be between -1 and 2,147,483,647 - -1 will result with 200 records as this is the default record cap.

### Not Required

==================

Argument 9: Filter for the IDO

Argument 10: OrderBy for the IDO
