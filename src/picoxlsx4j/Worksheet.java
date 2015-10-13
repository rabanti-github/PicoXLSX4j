/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j;

import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import picoxlsx4j.Worksheet.CellDirection;
import picoxlsx4j.exception.FormatException;
import picoxlsx4j.exception.OutOfRangeException;
import picoxlsx4j.exception.UndefinedStyleException;
import picoxlsx4j.style.Style;

/**
 * Class representing the style sheet of a workbook
 */
public class Worksheet {
    
    /**
    * Default column width as constant
    */
    public static final float DEFAULT_COLUMN_WIDTH = 10f;
    
    /**
    * Default row height as constant
    */
    public static final float DEFAULT_ROW_HEIGHT = 15f;
    
    /**
    * Enum to define the direction when using AddNextCell method
    */
    public enum CellDirection
    {
        /**
         * The next cell will be on the same row (A1,B1,C1...)
         */
        ColumnToColum,
        /**
         * The next cell will be on the same column (A1,A2,A3...)
         */
        RowToRow
    }
    
    private CellDirection currentCellDirection;
    private Style activeStyle;
    private Workbook workbookReference;
    private String sheetName;
    private int currentRowNumber;
    private int currentColumnNumber;
    private Map<String, Cell> cells;  
    private float defaultRowHeight;
    private float defaultColumnWidth;
    private Map<Integer, Float> columnWidths;
    private Map<Integer, Float> rowHeights;
    private int sheetID;
    
    public String getSheetName() {
        return sheetName;
    }

    /**
     * Sets the name of the sheet
     * @param sheetName Name of the sheet
     * @throws FormatException Thrown if the name contains illegal characters or is longer than 31 characters
     */
    public void setSheetName(String sheetName) throws FormatException {
        if (Helper.isNullOrEmpty(sheetName))
        {
            throw new FormatException("The sheet name must be between 1 and 31 characters");
        }
        if (sheetName.length() > 31)
        {
            throw new FormatException("The sheet name must be between 1 and 31 characters");
        }
        Pattern pattern = Pattern.compile("[\\[\\]\\*\\?/\\\\]");
        Matcher mx = pattern.matcher(sheetName);
        if (mx.groupCount() > 0)
        {
            throw new FormatException("The sheet name must must not contain the characters [  ]  * ? / \\ ");
        }
        this.sheetName = sheetName;
    }    

    /**
     * Gets the internal ID of the worksheet
     * @return Worksheet ID
     */
    public int getSheetID() {
        return sheetID;
    }
    /**
     * Sets the internal ID of the worksheet
     * @param sheetID Worksheet ID
     */
    public void setSheetID(int sheetID) {
        this.sheetID = sheetID;
    }
    
    /**
     * Gets the cells as list of the worksheet
     * @return List of Cell objects
     */
    public Map<String, Cell> getCells() {
        return cells;
    }    

    /*
     * Gets the default Row height
     * @return Default Row height
     */
    public float getDefaultRowHeight() {
        return defaultRowHeight;
    }
    
    /**
     * Sets the default Row height
     * @param defaultRowHeight Default Row height
     */
    public void setDefaultRowHeight(float defaultRowHeight) {
        this.defaultRowHeight = defaultRowHeight;
    }

    /**
     * Gets the default column width
     * @return Default column width
     */
    public float getDefaultColumnWidth() {
        return defaultColumnWidth;
    }

    /**
     * Sets the default column width
     * @param defaultColumnWidth Default column width
     */
    public void setDefaultColumnWidth(float defaultColumnWidth) {
        this.defaultColumnWidth = defaultColumnWidth;
    }

    /**
     * Gets the map of column widths. Key is the column number (zero-based), value is a float from 0 to 255.0
     * @return Map of column widths
     */
    public Map<Integer, Float> getColumnWidths() {
        return columnWidths;
    }

    /**
     * Gets the map of row heights. Key is the row number (zero-based), value is a float from 0 to 409.5
     * @return Map of row heights
     */
    public Map<Integer, Float> getRowHeights() {
        return rowHeights;
    }
    
    /**
     * Gets the direction when using AddNextCell method
     * @return Cell direction
     */
    public CellDirection getCurrentCellDirection() {
        return currentCellDirection;
    }

    /**
     * Sets the direction when using AddNextCell method
     * @param currentCellDirection Cell direction
     */
    public void setCurrentCellDirection(CellDirection currentCellDirection) {
        this.currentCellDirection = currentCellDirection;
    }

    /**
     * Default constructor
     */
    public Worksheet()
    {
        init();
    }
    
    /**
     * Constructor with name and sheet ID
     * @param name Name of the worksheet
     * @param id ID of the worksheet (for internal use)
     * @throws FormatException Thrown if the name contains illegal characters or is to long
     */
    public Worksheet(String name, int id) throws FormatException
    {
        init();
        setSheetName(name);
        this.sheetID = id;
    }    

    /**
     * Init method for constructors
     */
    private void init()
    {
        this.currentCellDirection = CellDirection.ColumnToColum;
        this.cells = new HashMap<String, Cell>();
        this.currentRowNumber = 0;
        this.currentColumnNumber = 0;
        this.defaultColumnWidth = DEFAULT_COLUMN_WIDTH;
        this.defaultRowHeight = DEFAULT_ROW_HEIGHT;
        this.columnWidths = new HashMap<>();
        this.rowHeights = new HashMap<>();
        this.activeStyle = null;
        this.workbookReference = null;        
    }
    
/* ************************************************* */ 
    /**
     * Adds a object to the next cell position. If the type of the value does not match with one of the supported data types, it will be casted to a String
     * @param value Unspecified value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCell(Object value) throws OutOfRangeException, UndefinedStyleException
    {
        Cell c = new Cell(value, Cell.CellType.DEFAULT, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    } 
    
    /**
     * Adds a string value to the next cell position
     * @param value String value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCell(String value) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.STRING, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    } 
    
    /**
     * Adds a integer value to the next cell position
     * @param value Integer value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addNextCell(int value) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.NUMBER, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    }    
    
    /**
     * Adds a double value to the next cell position
     * @param value Double value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCell(double value) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.NUMBER, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    }
    
    /**
     * Adds a float value to the next cell position
     * @param value Float value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addNextCell(float value) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.NUMBER, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    } 
    
    /**
     * Adds a date value to the next cell position
     * @param value Date value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addNextCell(Date value) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.DATE, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    }    

    /**
     * Adds a boolean value to the next cell position
     * @param value Boolean value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addNextCell(boolean value) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.BOOL, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    }
    
    /**
     * Adds a formula as string to the next cell position
     * @param formula Formula to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addNextCellFormula(String formula) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    }
    
    /**
     * Method to insert a generic cell to the next cell position
     * @param cell Cell object to insert
     * @param increment If true, the address value (row or column) will be incremented, otherwise not
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    private void addNextCell(Cell cell, boolean increment) throws UndefinedStyleException, OutOfRangeException
    {
        if (this.activeStyle != null)
        {
            cell.setStyle(this.activeStyle, this.workbookReference);
        }
        String address = cell.getCellAddressString();
        this.cells.put(address, cell);
        if (increment == true)
        {
            if (this.getCurrentCellDirection() == CellDirection.ColumnToColum)
            {
                this.currentColumnNumber++;
            }
            else
            {
                this.currentRowNumber++;
            }
        }
        else
        {
            if (this.getCurrentCellDirection() == CellDirection.ColumnToColum)
            {
                this.currentColumnNumber = cell.getColumnAddress() + 1;
                this.currentRowNumber = cell.getRowAddress();
            }
            else
            {
                this.currentColumnNumber = cell.getColumnAddress();
                this.currentRowNumber = cell.getRowAddress() + 1;
            }
        }        
    }
/* ************************************************* */  
    
    /**
     * Adds a object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String
     * @param value Unspecified value to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, int columnAddress, int rowAddress) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.DEFAULT, columnAddress, rowAddress);
        addNextCell(c, false);
    }
    
    /**
     * Adds a object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String
     * @param value Unspecified value to insert
     * @param address Cell address in the format A1 - XFD16384
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, String address) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row);
    }  
    
    /**
     * Adds a string value to the defined cell address
     * @param value String value to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(String value, int columnAddress, int rowAddress) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.STRING, columnAddress, rowAddress);
        addNextCell(c, false);
    }  
    
    /**
     * Adds a string value to the defined cell address
     * @param value String value to insert
     * @param address Cell address in the format A1 - XFD16384
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(String value, String address) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row);
    }  
    
    /**
     * Adds a integer value to the defined cell address
     * @param value Integer value to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(int value, int columnAddress, int rowAddress) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.NUMBER, columnAddress, rowAddress);
        addNextCell(c, false);
    }
    
    /**
     * Adds a integer value to the defined cell address
     * @param value Integer value to insert
     * @param address Cell address in the format A1 - XFD16384
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addCell(int value, String address) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row);
    }    
    
    /**
     * Adds a double value to the defined cell address
     * @param value Double value to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(double value, int columnAddress, int rowAddress) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.NUMBER, columnAddress, rowAddress);
        addNextCell(c, false);
    } 
    
    /**
     * Adds a double value to the defined cell address
     * @param value Double value to insert
     * @param address Cell address in the format A1 - XFD16384
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addCell(double value, String address) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row);
    }  
    
    /**
     * Adds a float value to the defined cell address
     * @param value Float value to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(float value, int columnAddress, int rowAddress) throws UndefinedStyleException, OutOfRangeException
   {
       Cell c = new Cell(value, Cell.CellType.NUMBER, columnAddress, rowAddress);
       addNextCell(c, false);
   } 
    
    /**
     * Adds a float value to the defined cell address
     * @param value Float value to insert
     * @param address Cell address in the format A1 - XFD16384
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addCell(float value, String address) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row);
    }  
    
    /**
     * Adds a date value to the defined cell address
     * @param value Date value to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Date value, int columnAddress, int rowAddress) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.DATE, columnAddress, rowAddress);
        addNextCell(c, false);
    }    
    
    /**
     * Adds a date value to the defined cell address
     * @param value Date value to insert
     * @param address Cell address in the format A1 - XFD16384
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addCell(Date value, String address) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row);
    } 
    
    /**
     * Adds a boolean value to the defined cell address
     * @param value Boolean value to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void AddCell(boolean value, int columnAddress, int rowAddress) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(value, Cell.CellType.BOOL, columnAddress, rowAddress);
        addNextCell(c, false);
    } 
    
     /**
     * Adds a boolean value to the defined cell address
     * @param value Boolean value to insert
     * @param address Cell address in the format A1 - XFD16384
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(boolean value, String address) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row);
    }     
/* ************************************************* */      
 
    /**
     * Adds a cell formula as string to the defined cell address
     * @param formula Formula to insert
     * @param address Cell address in the format A1 - XFD16384
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellFormula(String formula, String address) throws UndefinedStyleException, FormatException, OutOfRangeException
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        Cell c = new Cell(formula, Cell.CellType.FORMULA, adr.Column, adr.Row);
        addNextCell(c, false);
    }
    
    /**
     * Adds a cell formula as string to the defined cell address
     * @param formula Formula to insert
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void AddCellFormula(String formula, int columnAddress, int rowAddress) throws UndefinedStyleException, OutOfRangeException
    {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, columnAddress, rowAddress);
        addNextCell(c, false);
    }    
    
/* AddCellRange ************************************************* */     
    
    /**
     * Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs distinct function name.
     * @param values List of unspecified objects to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void AddObjectCellRange(List<Object> values, Cell.Address startAddress, Cell.Address endAddress) throws OutOfRangeException, UndefinedStyleException
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }    

    /**
     * Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of unspecified objects to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addObjectCellRange(List<Object> values, String cellRange) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    }    
    
    /**
     * Adds a list of string values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs distinct function name.
     * @param values List of string values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addStringCellRange(List<String> values, Cell.Address startAddress, Cell.Address endAddress) throws OutOfRangeException, UndefinedStyleException
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }    
    
    /**
     * Adds a list of string values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of string values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addStringCellRange(List<String> values, String cellRange) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    } 
    
    /**
     * Adds a list of integer values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs distinct function name.
     * @param values List of integer values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */       
    public void addIntegerCellRange(List<Integer> values, Cell.Address startAddress, Cell.Address endAddress) throws OutOfRangeException, UndefinedStyleException
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }    
    
    /**
     * Adds a list of integer values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of integer values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addIntegerCellRange(List<Integer> values, String cellRange) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    }     
    
     /**
     * Adds a list of double values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs distinct function name.
     * @param values List of double values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */   
    public void addDoubleCellRange(List<Double> values, Cell.Address startAddress, Cell.Address endAddress) throws OutOfRangeException, UndefinedStyleException
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }    

    /**
     * Adds a list of double values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of double values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addDoubleCellRange(List<Double> values, String cellRange) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    } 
    
    /**
     * Adds a list of float values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs distinct function name.
     * @param values List of float values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */       
    public void addFloatCellRange(List<Float> values, Cell.Address startAddress, Cell.Address endAddress) throws OutOfRangeException, UndefinedStyleException
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }  
    
    /**
     * Adds a list of float values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of float values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addFloatCellRange(List<Float> values, String cellRange) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    }   
    
    /**
     * Adds a list of date values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs distinct function name.
     * @param values List of date values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */       
    public void addDateCellRange(List<Date> values, Cell.Address startAddress, Cell.Address endAddress) throws OutOfRangeException, UndefinedStyleException
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }  
    
    /**
     * Adds a list of date values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of date values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addDateCellRange(List<Date> values, String cellRange) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    }    
 
    /**
     * Adds a list of boolean values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs distinct function name.
     * @param values List of boolean values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */        
    public void addBooleanCellRange(List<Boolean> values, Cell.Address startAddress, Cell.Address endAddress) throws OutOfRangeException, UndefinedStyleException
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }  
    
    /**
     * Adds a list of date values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of date values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addBooleanCellRange(List<Boolean> values, String cellRange) throws FormatException, OutOfRangeException, UndefinedStyleException
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    }    
    
    /**
     * Internal function to add a generic list of value to the defined cell range
     * @param <T> Data type of the generic value list
     * @param values List of values
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws OutOfRangeException Thrown if the next cell is out of range (on row or column)
     */
    private <T> void addCellRangeInternal(List<T> values, Cell.Address startAddress, Cell.Address endAddress) throws OutOfRangeException, UndefinedStyleException
    {
        List<Cell.Address> addresses = Cell.getCellRange(startAddress, endAddress);
        if (values.size() != addresses.size())
        {
            throw new OutOfRangeException("The number of passed values (" + Integer.toString(values.size()) + ") differs from the number of cells within the range (" + Integer.toString(addresses.size()) + ")");
        }
        List<Cell> list = Cell.convertArray(values);
        int len = values.size();
        for(int i = 0; i < len; i++)
        {
            list.get(i).setRowAddress(addresses.get(i).Row);
            list.get(i).setColumnAddress(addresses.get(i).Column);
            addNextCell(list.get(i), false);
        }
    }
    
/* RemoveCell ************************************************* */  
    
    /**
     * Removes a previous inserted cell at the defined address
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @return Returns true if the cell could be removed (existed), otherwise false (did not exist)
     * @throws OutOfRangeException Thrown if the resolved cell address is out of range
     */
    public boolean removeCell(int columnAddress, int rowAddress) throws OutOfRangeException
    {
        String address = Cell.resolveCellAddress(columnAddress, rowAddress);
        if (this.cells.containsKey(address))
        {
            this.cells.remove(address);
            return true;
        }
        else
        {
            return false;
        }
    }   
    
    /**
     * Removes a previous inserted cell at the defined address
     * @param address Cell address in the format A1 - XFD16384
     * @return Returns true if the cell could be removed (existed), otherwise false (did not exist)
     * @throws OutOfRangeException Thrown if the resolved cell address is out of range
     * @throws FormatException Thrown if the passed address is malformed
     */
    public boolean removeCell(String address) throws OutOfRangeException, FormatException
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        return removeCell(adr.Column, adr.Row);
    } 
/* ************************************************* */     
  
    /**
     * Moves the current position to the next column
     */
    public void goToNextColumn()
    {
        this.currentColumnNumber++;
        this.currentRowNumber = 0;
    }    
    
    /**
     * Moves the current position to the next row (use for a new line)
     */
    public void goToNextRow()
    {
        this.currentRowNumber++;
        this.currentColumnNumber = 0;
    }    
    
    /**
     * Sets the current row address (row number, zero based)
     * @param rowAddress Row number (zero based)
     * @throws OutOfRangeException Thrown if the address is out of the valid range. Range is from 0 to 1048575 (1048576 rows)
     */
    public void setCurrentRowAddress(int rowAddress) throws OutOfRangeException
    {
        if (rowAddress >= 1048576 || rowAddress < 0)
        {
            throw new OutOfRangeException("The row number (" + Integer.toString(rowAddress) + ") is out of range. Range is from 0 to 1048575 (1048576 rows).");
        }
        this.currentRowNumber = rowAddress;
    }
    
    /**
     * Sets the current column address (column number, zero based)
     * @param columnAddress Column number (zero based)
     * @throws OutOfRangeException Thrown if the address is out of the valid range. Range is from 0 to 16383 (16384 columns)
     */
    public void setCurrentColumnAddress(int columnAddress) throws OutOfRangeException
    {
        if (columnAddress >= 16383 || columnAddress < 0)
        {
            throw new OutOfRangeException("The column number (" + Integer.toString(columnAddress) + ") is out of range. Range is from 0 to 16383 (16384 columns).");
        }
        this.currentColumnNumber = columnAddress;
    }   
    
    /**
     * Set the current cell address
     * @param address Cell address in the format A1 - XFD16384
     * @throws OutOfRangeException Thrown if the address is out of the valid range. Range is for rows from 0 to 1048575 (1048576 rows) and for columns from 0 to 16383 (16384 columns)
     * @throws FormatException Thrown if the passed address is malformed
     */
    public void setCurentCellAddress(String address) throws OutOfRangeException, FormatException
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        setCurentCellAddress(adr.Column, adr.Row);
    }  
    
    /**
     * Set the current cell address
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws OutOfRangeException Thrown if the address is out of the valid range. Range is for rows from 0 to 1048575 (1048576 rows) and for columns from 0 to 16383 (16384 columns)
     */
    public void setCurentCellAddress(int columnAddress, int rowAddress) throws OutOfRangeException
    {
        setCurrentColumnAddress(columnAddress);
        setCurrentRowAddress(rowAddress);
    }    
    
    /**
     * Set the current cell address
     * @param columnAddress Column number (zero based)
     * @param width Row number (zero based)
     * @throws OutOfRangeException Thrown if the address is out of the valid range. Range is from 0 to 16383 (16384 columns)
     */
    public void SetColumnWidth(String columnAddress, float width) throws OutOfRangeException
    {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnWidth(columnNumber, width);
    }
    
    /**
     * Sets the width of the passed column number (zero-based)
     * @param columnNumber Column number (zero-based, from 0 to 16383)
     * @param width Width from 0 to 255.0
     * @throws OutOfRangeException Thrown if the address is out of the valid range. Range is from 0 to 16383 (16384 columns)
     */
    public void setColumnWidth(int columnNumber, float width) throws OutOfRangeException
    {
        if (columnNumber >= 16384 || columnNumber < 0)
        {
            throw new OutOfRangeException("The column number (" + Integer.toString(columnNumber) + ") is out of range. Range is from 0 to 16383 (16384 columns).");
        }
        if (width < 0 || width > 255)
        {
            throw new OutOfRangeException("The column width (" + Float.toString(width) + ") is out of range. Range is from 0 to 255 (chars).");
        }
        this.columnWidths.put(columnNumber, width);
    }  
   
    /**
     * Sets the height of the passed row number (zero-based)
     * @param rowNumber Row number (zero-based, 0 to 1048575)
     * @param height Height from 0 to 409.5
     * @throws OutOfRangeException Thrown if the address is out of the valid range. Range is from 0 to 1048575 (1048576 rows)
     */
    public void setRowHeight(int rowNumber, float height) throws OutOfRangeException
   {
       if (rowNumber >= 1048576 || rowNumber < 0)
       {
           throw new OutOfRangeException("The row number (" + Integer.toString(rowNumber) + ") is out of range. Range is from 0 to 1048575 (1048576 rows).");
       }
       if (height < 0 || height > 409.5)
       {
           throw new OutOfRangeException("The row height (" + Float.toString(height) + ") is out of range. Range is from 0 to 409.5 (equals 546px).");
       }
       this.rowHeights.put(rowNumber, height);
   }  
    
    /**
     * Sets the active style of the worksheet. This style will be assigned to all later added cells
     * @param style Style to set as active style
     * @param workbookReference Reference to the workbook. All stiles are managed in this workbook
     */
    public void setActiveStyle(Style style, Workbook workbookReference)
    {
        this.activeStyle = style;
        this.workbookReference = workbookReference;
    } 
    
    /**
     * Clears the active style of the worksheet. All later added calls will contain no style unless another active style is set
     */
    public void clearActiveStyle()
    {
        this.activeStyle = null;
        this.workbookReference = null;
    }    
    
    
}
