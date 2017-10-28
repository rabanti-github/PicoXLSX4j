# PicoXLSX4j
PicoXLSX4j is a small Java library to create XLSX files (Microsoft Excel 2007 or newer) in an easy and native way. It is a direct port of [PicoXLSX for C#](https://github.com/rabanti-github/PicoXLSX)

* No need for an installation of Microsoft Office
* No need for Office interop/DCOM or other bridging libraries
* No need for 3rd party libraries
* Pure usage of standard JRE

# What's new in version 2.x
* Changed package structure to ch.rabanti.picoxlsx4j
* Complete replacement of the old style handling
* Added more options to assign styles to cells
* Added an option for sanitizing of worksheet names
* Replaced specific exception classes with general exceptions (e.g. StyleException, FormatException or WorksheetException)
* Added functions to retrieve stored data and the current cell address
* Many internal optimizations and additional documentation

# Requirements
PicoXLSX4j was created with Java 8 and is fully compatible with Java 7<br>
The only requirement for developments are a current JDK to develop and JRE to run.

# Installation
## As JAR
Simply place the PicoXLSX4j.jar into the lib folder of your project and create a library reference to it in your IDE.
## As source files
Place all .java files from the PicoXLSX4j source folder into your project. The folder structure defines the packages. Please use refactoring if you want to relocate the files.

# Usage
## Quick Start
```
 Workbook workbook = new Workbook("myWorkbook.xlsx", "Sheet1");       // Create new workbook with a worksheet called Sheet1
 workbook.getCurrentWorksheet().addNextCell("Some Data");             // Add cell A1
 workbook.getCurrentWorksheet().addNextCell(42);                      // Add cell B1
 workbook.getCurrentWorksheet().goToNextRow();                        // Go to row 2
 workbook.getCurrentWorksheet().addNextCell(new Date());              // Add cell A2
 workbook.Save();                                                     // Save the workbook as myWorkbook.xlsx
```

## Further References
See the full <b>API-Documentation</b> at: [https://rabanti-github.github.io/PicoXLSX4j/](https://rabanti-github.github.io/PicoXLSX4j/).<br>
The [Demo class](https://github.com/rabanti-github/PicoXLSX4j/blob/master/src/ch/rabanti/picoxlsx4j/demo/PicoXLSX4j.java) contains nine simple use cases. You can find also the full documentation in the [Javadoc-Folder](https://github.com/rabanti-github/PicoXLSX4j/tree/master/dist/javadoc) or as Javadoc annotations in the .java files.<br>
