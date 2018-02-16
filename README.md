# PicoXLSX4j
![PicoXLSX](https://rabanti-github.github.io/PicoXLSX/icons/PicoXLSX.png)


PicoXLSX4j is a small Java library to create XLSX files (Microsoft Excel 2007 or newer) in an easy and native way. It is a direct port of [PicoXLSX for C#](https://github.com/rabanti-github/PicoXLSX)

* No need for an installation of Microsoft Office
* No need for Office interop/DCOM or other bridging libraries
* No need for 3rd party libraries
* Pure usage of standard JRE

Project website: [https://picoxlsx.rabanti.ch](https://picoxlsx.rabanti.ch)

See the **[Change Log](https://github.com/rabanti-github/PicoXLSX4j/blob/master/Changelog.md)** for recent updates.

# What's new in version 2.x
* Changed package structure to ch.rabanti.picoxlsx4j
* Added maven support (available in maven central under ch.rabanti.picoxlsx4j)
* Complete replacement of the old style handling
* Added more options to assign styles to cells
* Added Shortner (property WS) to reduce the code overhead
* Added static methods for the most important formulas (round, floor, ceil, min, max, average, median, sum, vlookup)
* Added Save option to save the XLSX file as stream
* Added an option for sanitizing of worksheet names
* Replaced specific exception classes with general exceptions (e.g. StyleException, FormatException or WorksheetException)
* Added functions to retrieve stored data and the current cell address
* Many internal optimizations and additional documentation

# Requirements
PicoXLSX4j was created with Java 8 and is fully compatible with Java 7<br>
The only requirement for developments are a current JDK to develop and JRE to run.

# Installation
## As JAR
Simply place the PicoXLSX4j jar file (e.g. **picoxlsx4j-2.3.1.jar**) into the lib folder of your project and create a library reference to it in your IDE.
## As source files
Place all .java files from the PicoXLSX4j source folder into your project. The folder structure defines the packages. Please use refactoring if you want to relocate the files.
## Maven ##
**Note - February 16, 2018: The library was released to OSSRH but not yet synchronized with Maven Central. This section will be updated as soon as the artifact is publicly available.**
Add the following information to your POM file:
```xml
<dependency>
    <groupId>ch.rabanti</groupId>
    <artifactId>picoxlsx4j</artifactId>
    <version>2.3.1</version>
</dependency>
``` 

**Important:** The version number may change. Please check the [Change Log](https://github.com/rabanti-github/PicoXLSX4j/blob/master/Changelog.md) for the most recent version. The keywords ```LATEST```  and ```RELEASE``` are only valid in Maven 2, not 3 and newer. 


# Usage
## Quick Start (shortened syntax)
```java
 Workbook workbook = new Workbook("myWorkbook.xlsx", "Sheet1");         // Create new workbook with a worksheet called Sheet1
 workbook.WS.value("Some Data");                                        // Add cell A1
 workbook.WS.formula("=A1");                                            // Add formula to cell B1
 workbook.WS.down();                                                    // Go to row 2
 workbook.WS.value(new Date(), BasicStyles.Bold());                     // Add formated value to cell A2
 try{
   workbook.save();                                                     // Save the workbook as myWorkbook.xlsx
 } catch (Exception ex) {}
```

## Quick Start (regular syntax)
```java
 Workbook workbook = new Workbook("myWorkbook.xlsx", "Sheet1");       // Create new workbook with a worksheet called Sheet1
 workbook.getCurrentWorksheet().addNextCell("Some Data");             // Add cell A1
 workbook.getCurrentWorksheet().addNextCell(42);                      // Add cell B1
 workbook.getCurrentWorksheet().goToNextRow();                        // Go to row 2
 workbook.getCurrentWorksheet().addNextCell(new Date());              // Add cell A2
 try {
   workbook.Save();                                                   // Save the workbook as myWorkbook.xlsx
 } catch (Exception ex) {}
```

## Further References
See the full <b>API-Documentation</b> at: [https://rabanti-github.github.io/PicoXLSX4j/](https://rabanti-github.github.io/PicoXLSX4j/).<br>
The [Demo class](https://github.com/rabanti-github/PicoXLSX4j/blob/master/src/ch/rabanti/picoxlsx4j/demo/PicoXLSX4j.java) contains twelve simple use cases. You can find also the full documentation in the [Javadoc-Folder](https://github.com/rabanti-github/PicoXLSX4j/tree/master/dist/javadoc) or as Javadoc annotations in the .java files.<br>
