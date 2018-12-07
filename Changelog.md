# Change Log


## v2.5.4

---
Release Date: **07.12.2018**

- Improved the performance of adding stylized cells by factor 10 to 100
- Code reformatting



## v2.5.3

---
Release Date: **04.11.2018**

- Fixed a bug in the style handling of merged cells. Bug fix provided by David Courtel for PicoXLSX (C#)


## v2.5.2

---
Release Date: **24.08.2018**

- Fixed a bug in the calculation of OA Dates (internal format)
- Documentation Update


## v2.5.1

---
Release Date: **19.08.2018**

- Fixed a bug in the Font style class
- Fixed typos


## v2.5.0

---
Release Date: **03.07.2018**

- Added address types (no fixed rows and columns, fixed rows, fixed columns, fixed rows and columns)
- Added new CellDirection Disabled, if the addresses of the cells are defined manually (addNextCell will override the current cell in this case)
- Altered Demo 3 to demonstrate disabling of automatic cell addressing
- Extended Demo 1 to demonstrate the new address types
- Minor, internal changes


## v2.4.0

---
Release Date: **08.06.2018**

- Added style appending (builder / method chaining)
- Added new basic styles colorizedText, colorizedBackground and font as functions
- Added a new constructor for Workbooks without file name to handle stream-only workbooks more logical
- Added the functions hasCell, getLastColumnNumber and getLastRowNumber in the Worksheet class
- Renamed the function SetColor in the class Fill (Style) to setColor, to follow conventions. Minor refactoring in existing projects may be possible
- Fixed a bug when overriding a worksheet name with sanitizing
- Added new demo for the introduced style features
- Internal optimizations and fixes


## v2.3.4

---
Release Date: **31.05.2018**

- Fixed a bug in the processing of column widths. Bug fix provided by Johan Lindvall for PicoXLSX, adapted for PicoXLSX4j and NanoXLSX4j
- Added numeric data types Byte, BigDecimal, and Short (proposal by Johan Lindvall for PicoXLSX)
- Changed the behavior of cell type casting. User defined cell types will now only be overwritten if the type is DEFAULT (proposal by Johan Lindvall for PicoXLSX)


## v2.3.3

---
Release Date: **26.05.2018**

- Fixed a bug in the handling of worksheet protection
- Code cleanup


## v2.3.2
 
---
Release Date: **12.03.2018**

**Note**: Due to some refactoring (see below) in this version, changes of existing code may be necessary. However, most introduced changes are on a rather low level and probably only used internally although publicly accessible

- Renamed the getters and setters get/setRowAddress and get/setColumnAddress to get/setRowNumber and get/setColumnNumber in the class Cell for clarity
- Renamed the methods getCurrentColumnAddress, getCurrentRowAddress, setCurrentColumnAddress and setCurrentRowAddress in the class Worksheet to getCurrentColumnNumber, getCurrentRowNumber, setCurrentColumnNumber and SetCurrentRowNumber for clarity
- Renamed the constants MIN_ROW_ADDRESS, MAX_ROW_ADDRESS, MIN_COLUMN_ADDRESS, MAX_COLUMN_ADDRESS in the class Worksheet to MIN_ROW_NUMBER, MAX_ROW_NUMBER, MIN_COLUMN_NUMBER, MAX_COLUMN_NUMBER for clarity
- Fixed typos
- Documentation update


## v2.3.1
 
---
Release Date: **16.02.2018**

**Note**: The naming of the released jar has changed due to the new Maven deployment process. The version number is now part of the jar file name. This release has no functional changes to version 2.3.0

- Changed project structure to a maven template (package structure has not changed)
- Adapted demos to store all demo files in a particular folder (keeps the root folder of the project clean)
- Changed the naming of the dist files. executable and javadoc (jar files) contains now the version number
- Removed the javadoc folder in dist in favour of a jar file (the plain files are still available in the docs folder)


## v2.3.0
 
---
Release Date: **13.02.2018**

- Added most important formulas as static method calls in the class BasicFormulas (round, floor, ceil, min, max, average, median, sum, vlookup)
- Removed overloaded methods to add cells as type Cell. This can be done now with the overloading of the type object (no code changes necessary)
- Added new constructors for Address and Range
- Added demo for the new static formula methods
- Changed the method names SetColumnWidth (Worksheet) to setColumnWidth and GetCellRange (Cell) to getCellRange, to follow conventions (minor updates of code may necessary)
- Minor bug fixes
- Documentation update
 

## v2.2.0

---
Release Date: **10.12.2017**

- Added Shortener class (WS) in workbook for quicker writing of data / formulas
- Documentation Update


## v2.1.0 

---
Release Date: **03.12.2017**

- Added saveToStream method in Workbook class
- Added demo for the new stream save method
- Changed log to MD format


## v2.0.1 

---
Release Date: **01.11.2017**

- Changed function name getCurrentColumnNumber to getCurrentColumnAddress (alignment to setter)
- Changed function name getCurrentRowNumber to getCurrentRowAddress (alignment to setter)
- Fixed errors in teh demo documentation (comments)
- Fixed typos


## v2.0.0 

---
Release Date: **29.10.2017**

**Note**: This major version is not compatible with code of v1.x. However, porting is feasible with moderate effort
- Changed package structure to ch.rabanti.picoxlsx4j
- Complete replacement of style handling
- Added an option to add styles with the cell values in one step
- Added a sanitizing function for worksheet names (with auto-sanitizing when adding a worksheet as option)
- Changed specific exceptions to general exceptions (e.g. StyleException, FormatException or WorksheetException)
- Added function to retrieve cell values easier
- Added functions to get the current column or row number
- Many internal optimizations
- Added more documentation
- Added new functionality to the demos

## v1.6.3 

---
Release Date: **24.08.2017**

- fixed a bug in the handling of border styles
- Added further null checks
- Minor optimizations
- Fixed typos

## v1.6.2 

---
Release Date: **12.08.2017**

- fixed a bug in the function to remove merged cells (Worksheet class)
- Fixed typos

## v1.6.1
 
---
Release Date: **08.08.2017**

**Note**: Due to a (now fixed) typo in a public parameter name, it is possible that some function calls on existing code must be fixed too (just renaming).
- Fixed typos (parameters and text)
- Minor optimization
- Moved Sub-Classes Range and Address to separate files
- Complete reformatting of code (alphabetical order)
- HTML documentation moved to folder 'docs' to provide an automatic API documentation on the hosting platform

## v1.6.0 

---
Release Date: **15.04.2017**

**Note**: Using this version of the library with old code can cause compatibility issues due to the simplification of some methods (see below).
- Simplified all style assignment methods. Referencing of the workbook is not necessary anymore (can cause compatibility issues with existing code; just remove the workbook references)
- Removed getCellAddressString and setCellAddressString  Method. Replaced by getCellAddress and setCellAddress (String object)
- getCellAddress and setCellAddress (Address object) are replaced by getCellAddress2 and setCellAddress2 
- Additional checks in the assignment methods for columns and rows
- Minor changes (code and documentation)

## v1.5.5

---
Release Date: **03.04.2017**

- Fixed a potential bug induced by non-Gregorian calendars (e.g Minguo, Heisei period, Hebrew) on the host system
- Code cleanup
- Minor bug fixes
- Fixed typos

## v1.5.4

---
Release Date: **20.03.2017**

- Extended the sanitizing of allowed XML characters according the XML specifications to avoid errors with illegal characters in passed strings
- Fixed typos

## v1.5.3

---
Release Date: **02.12.2016**

- Fixed bug in the handling of the cell types

## v1.5.2

---
Release Date: **17.11.2016**

- Fixed general bug in the handling of the sharedStrings table. Please update
- Passed null values to cells are now interpreted as empty values. Caused an exception until now

## v1.5.1

---
Release Date: **15.11.2016**

- Fixed bug in sharedStrings table


## v1.5.0

---
Release Date: **16.08.2016**

**Note**: Using this version of the library with old code can cause compatibility issues due to the removal of some methods (see below).
- Removed all overloaded methods with various input values for adding cells. Object is sufficient
- Added sharedStrings table to manage strings more efficient (Excel standard)
- Changed demos according to removed overloaded methods (ArrayList&lt;String&gt; is now ArrayList&lt;Object&gt;)
- Added support for long (64bit) data type

## v1.4.0

---
Release Date: **11.08.2016**

- Added support for cell selection
- Added support for worksheet selection
- Removed XML namespace 'x' as prefix in OOXML output. No use for this at the moment
- Removed newlines from OOXML output. No relevance for parser
- Added further demo for the new features

## v1.3.0  

---
Release Date: **18.01.2016**

- Added support for auto filter (columns)
- Added support for hiding columns and rows
- Added new Column class (sub-class of Worksheet) to manage column based properties more efficiently
- Removed unused Exception UnsupportedDataTypeException
- Fixed some documentation issues
- Minor bug fixes + typos 
- Added further demo for the new features

## v1.2.4	

---
Release Date: **08.11.2015**

- Fixed a bug in the meta data section

## v1.2.3	

---
Release Date: **04.11.2015**

- Changed all Exceptions to RuntimeExceptions, apart from IOException
- Fixed typos

## v1.2.2	

---
Release Date: **02.11.2015**

- Added support for protecting workbooks

## v1.2.1	

---
Release Date: **01.11.2015**

- Added support to protect worksheets with a password
- Minor bug fixes
- Added more documentation

## v1.2.0	

---
Release Date: **31.10.2015**

- Added support for merging cells
- Added support for Protecting worksheets (no support for passwords yet)
- Minor bug fixes
- Fixed typos
- Added further demo for the new features

## v1.1.3
	
---
Release Date: **13.10.2015**

- Update to Java 8 (full compatible to Java 7)
- Fixed Javadoc to meet requirements of Java 8

## v1.1.2	

---
Release Date: **12.10.2015**

- Initial release (synced to v1.1.2 of PicoXLSX for C#)

