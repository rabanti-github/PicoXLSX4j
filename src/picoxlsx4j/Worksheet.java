/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2016
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j;

import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import picoxlsx4j.Worksheet.CellDirection;
import picoxlsx4j.exception.FormatException;
import picoxlsx4j.exception.OutOfRangeException;
import picoxlsx4j.exception.UnknownRangeException;
import picoxlsx4j.exception.UndefinedStyleException;
import picoxlsx4j.style.Style;

/**
 * Class representing a worksheet of a workbook
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
    

    /**
     * Enum to define the possible protection types when protecting a worksheet
     */
        public enum SheetProtectionValue
        {
           // sheet, // Is alway on 1 if protected
            /**
            If selected, the user can edit objects if the worksheets is protected
            */
            objects,
            /**
            If selected, the user can edit scenarios if the worksheets is protected
            */
            scenarios,
            /**
            If selected, the user can format cells if the worksheets is protected
            */
            formatCells,
            /**
            If selected, the user can format columns if the worksheets is protected
            */
            formatColumns,
            /**
            If selected, the user can format rows if the worksheets is protected
            */
            formatRows,
            /**
            If selected, the user can insert columns if the worksheets is protected
            */
            insertColumns,
            /**
            If selected, the user can insert rows if the worksheets is protected
            */
            insertRows,
            /**
            If selected, the user can insert hyperlinks if the worksheets is protected
            */
            insertHyperlinks,
            /**
            If selected, the user can delete columns if the worksheets is protected
            */
            deleteColumns,
            /**
            If selected, the user can delete rows if the worksheets is protected
            */
            deleteRows,
            /**
            If selected, the user can select locked cells if the worksheets is protected
            */
            selectLockedCells,
            /**
            If selected, the user can sort cells if the worksheets is protected
            */
            sort,
            /**
            If selected, the user can use auto filters if the worksheets is protected
            */
            autoFilter,
            /**
            If selected, the user can use pivot tables if the worksheets is protected
            */
            pivotTables,
            /**
            If selected, the user can select unlocked cells if the worksheets is protected
            */
            selectUnlockedCells 
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
    //private Map<Integer, Float> columnWidths;
    private Map<Integer, Float> rowHeights;
    private Map<String, Cell.Range> mergedCells;
    private Map<Integer, Column> columns;
    private Map<Integer, Boolean> hiddenRows;
    private boolean useSheetProtection;
    private List<SheetProtectionValue> sheetProtectionValues;
    private String sheetProtectionPassword;
    private int sheetID;
    private Cell.Range autoFilterRange;
    private Cell.Range selectedCells;

    /**
     * Gets the range of selected cells of this worksheet. Null if no cells are selected
     * @return Cell range of the selected cells
     */
    public Cell.Range getSelectedCells() {
        return selectedCells;
    }
    
    /**
     * Gets the name of the sheet
     * @return Name of the sheet
     */
    public String getSheetName() {
        return sheetName;
    }

    /**
     * Sets the name of the sheet
     * @param sheetName Name of the sheet
     * @throws FormatException Thrown if the name contains illegal characters or is longer than 31 characters
     */
    public void setSheetName(String sheetName) {
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
/*
    / **
     * Gets the map of column widths. Key is the column number (zero-based), value is a float from 0 to 255.0
     * @return Map of column widths
     * /
    public Map<Integer, Float> getColumnWidths() {
        return columnWidths;
    }
*/
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
     * Gets whether the worksheet is protected
     * @return If true, the worksheet is protected
     */
    public boolean isUseSheetProtection() {
        return useSheetProtection;
    }

    /**
     * Sets whether the worksheet is protected
     * @param useSheetProtection If true, the worksheet is protected
     */
    public void setUseSheetProtection(boolean useSheetProtection) {
        this.useSheetProtection = useSheetProtection;
    }

    /**
     * Gets the list of SheetProtectionValue. These values defines the allowed actions if the worksheet is protected
     * @return List of SheetProtectionValues
     */
    public List<SheetProtectionValue> getSheetProtectionValues() {
        return sheetProtectionValues;
    }  

    /**
     * Gets the map with merged cells (only references)
     * @return Hashmap with merged cell references
     */
    public Map<String, Cell.Range> getMergedCells() {
        return mergedCells;
    }

    /**
     * Gets the password used for sheet protection
     * @return Password (UTF-8)
     */
    public String getSheetProtectionPassword() {
        return sheetProtectionPassword;
    } 

    /**
     * Gets the map of all columns with non-standard properties, like auto filter applied or a special width
     * @return map of columns
     */
    public Map<Integer, Column> getColumns() {
        return columns;
    }

    /**
     * Gets the range of the auto filter. If null, no auto filters are applied
     * @return Range of auto filter
     */
    public Cell.Range getAutoFilterRange() {
        return autoFilterRange;
    }

    /**
     * Gets the Map of hidden rows.  Key is the row number (zero-based), value is a boolean. True indicates hidden, false visible.Entries with the value false are not affecting the worksheet. These entries can be removed<br>
     * @return Map with hidden rows
     */
    public Map<Integer, Boolean> getHiddenRows() {
        return hiddenRows;
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
    public Worksheet(String name, int id)
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
        this.rowHeights = new HashMap<>();
        this.activeStyle = null;
        this.workbookReference = null;
        this.mergedCells = new HashMap<>();
        this.sheetProtectionValues = new ArrayList<>();
        this.hiddenRows = new HashMap<>();
        this.columns = new HashMap<>();
    }
    
/* ************************************************* */ 
    /**
     * Adds a object to the next cell position. If the type of the value does not match with one of the supported data types, it will be casted to a String
     * @param value Unspecified value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCell(Object value)
    {
        Cell c = new Cell(value, Cell.CellType.DEFAULT, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    } 
    
    /**
     * Adds a string value to the next cell position
     * @param value String value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCell(String value)
    {
        Cell c = new Cell(value, Cell.CellType.STRING, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    } 
    
    /**
     * Adds a integer value to the next cell position
     * @param value Integer value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addNextCell(int value)
    {
        Cell c = new Cell(value, Cell.CellType.NUMBER, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    }    
    
    /**
     * Adds a double value to the next cell position
     * @param value Double value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addNextCell(double value)
    {
        Cell c = new Cell(value, Cell.CellType.NUMBER, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    }
    
    /**
     * Adds a float value to the next cell position
     * @param value Float value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addNextCell(float value)
    {
        Cell c = new Cell(value, Cell.CellType.NUMBER, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    } 
    
    /**
     * Adds a date value to the next cell position
     * @param value Date value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addNextCell(Date value)
    {
        Cell c = new Cell(value, Cell.CellType.DATE, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    }    

    /**
     * Adds a boolean value to the next cell position
     * @param value Boolean value to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addNextCell(boolean value)
    {
        Cell c = new Cell(value, Cell.CellType.BOOL, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    }
    
    /**
     * Adds a formula as string to the next cell position
     * @param formula Formula to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addNextCellFormula(String formula)
    {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, this.currentColumnNumber, this.currentRowNumber);
        addNextCell(c, true);
    }
    
    /**
     * Method to insert a generic cell to the next cell position
     * @param cell Cell object to insert
     * @param increment If true, the address value (row or column) will be incremented, otherwise not
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    private void addNextCell(Cell cell, boolean increment)
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
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, int columnAddress, int rowAddress)
    {
        Cell c = new Cell(value, Cell.CellType.DEFAULT, columnAddress, rowAddress);
        addNextCell(c, false);
    }
    
    /**
     * Adds a object to the defined cell address. If the type of the value does not match with one of the supported data types, it will be casted to a String
     * @param value Unspecified value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Object value, String address)
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
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(String value, int columnAddress, int rowAddress)
    {
        Cell c = new Cell(value, Cell.CellType.STRING, columnAddress, rowAddress);
        addNextCell(c, false);
    }  
    
    /**
     * Adds a string value to the defined cell address
     * @param value String value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(String value, String address)
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
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(int value, int columnAddress, int rowAddress)
    {
        Cell c = new Cell(value, Cell.CellType.NUMBER, columnAddress, rowAddress);
        addNextCell(c, false);
    }
    
    /**
     * Adds a integer value to the defined cell address
     * @param value Integer value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addCell(int value, String address)
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
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(double value, int columnAddress, int rowAddress)
    {
        Cell c = new Cell(value, Cell.CellType.NUMBER, columnAddress, rowAddress);
        addNextCell(c, false);
    } 
    
    /**
     * Adds a double value to the defined cell address
     * @param value Double value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addCell(double value, String address)
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
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(float value, int columnAddress, int rowAddress)
   {
       Cell c = new Cell(value, Cell.CellType.NUMBER, columnAddress, rowAddress);
       addNextCell(c, false);
   } 
    
    /**
     * Adds a float value to the defined cell address
     * @param value Float value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addCell(float value, String address)
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
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Date value, int columnAddress, int rowAddress)
    {
        Cell c = new Cell(value, Cell.CellType.DATE, columnAddress, rowAddress);
        addNextCell(c, false);
    }    
    
    /**
     * Adds a date value to the defined cell address
     * @param value Date value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addCell(Date value, String address)
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
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(boolean value, int columnAddress, int rowAddress)
    {
        Cell c = new Cell(value, Cell.CellType.BOOL, columnAddress, rowAddress);
        addNextCell(c, false);
    } 
    
     /**
     * Adds a boolean value to the defined cell address
     * @param value Boolean value to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(boolean value, String address)
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        addCell(value, adr.Column, adr.Row);
    }
    
    /**
     * Adds a cell object. This object must contain a valid row and column address
     * @param cell Cell object to insert
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCell(Cell cell)
    {
        addNextCell(cell, false);
    }    
    
/* addCellFormula ************************************************* */      
 
    /**
     * Adds a cell formula as string to the defined cell address
     * @param formula Formula to insert
     * @param address Cell address in the format A1 - XFD1048576
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellFormula(String formula, String address)
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
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addCellFormula(String formula, int columnAddress, int rowAddress)
    {
        Cell c = new Cell(formula, Cell.CellType.FORMULA, columnAddress, rowAddress);
        addNextCell(c, false);
    }    
    
/* addCellRange ************************************************* */     
    
    /**
     * Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs a distinct function name.
     * @param values List of unspecified objects to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addObjectCellRange(List<Object> values, Cell.Address startAddress, Cell.Address endAddress)
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }    

    /**
     * Adds a list of object values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of unspecified objects to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addObjectCellRange(List<Object> values, String cellRange)
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    }    
    
    /**
     * Adds a list of string values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs a distinct function name.
     * @param values List of string values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */    
    public void addStringCellRange(List<String> values, Cell.Address startAddress, Cell.Address endAddress)
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }    
    
    /**
     * Adds a list of string values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of string values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addStringCellRange(List<String> values, String cellRange)
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    } 
    
    /**
     * Adds a list of integer values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs a distinct function name.
     * @param values List of integer values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */       
    public void addIntegerCellRange(List<Integer> values, Cell.Address startAddress, Cell.Address endAddress)
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }    
    
    /**
     * Adds a list of integer values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of integer values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addIntegerCellRange(List<Integer> values, String cellRange)
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    }     
    
     /**
     * Adds a list of double values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs a distinct function name.
     * @param values List of double values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */   
    public void addDoubleCellRange(List<Double> values, Cell.Address startAddress, Cell.Address endAddress)
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }    

    /**
     * Adds a list of double values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of double values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addDoubleCellRange(List<Double> values, String cellRange)
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    } 
    
    /**
     * Adds a list of float values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs a distinct function name.
     * @param values List of float values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */       
    public void addFloatCellRange(List<Float> values, Cell.Address startAddress, Cell.Address endAddress)
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }  
    
    /**
     * Adds a list of float values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of float values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addFloatCellRange(List<Float> values, String cellRange)
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    }   
    
    /**
     * Adds a list of date values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs a distinct function name.
     * @param values List of date values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */       
    public void addDateCellRange(List<Date> values, Cell.Address startAddress, Cell.Address endAddress)
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }  
    
    /**
     * Adds a list of date values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of date values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addDateCellRange(List<Date> values, String cellRange)
    {
        Cell.Range rng = Cell.resolveCellRange(cellRange);
        addCellRangeInternal(values, rng.StartAddress, rng.EndAddress);
    }    
 
    /**
     * Adds a list of boolean values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String<br>
     * Note: Due to limitations of Java generics can this group of methods not be defined as overloading method with a single function name. Each inner type needs a distinct function name.
     * @param values List of boolean values to insert
     * @param startAddress Start address
     * @param endAddress End address
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */        
    public void addBooleanCellRange(List<Boolean> values, Cell.Address startAddress, Cell.Address endAddress)
    {
        addCellRangeInternal(values, startAddress, endAddress);
    }  
    
    /**
     * Adds a list of date values to a defined cell range. If the type of the a particular value does not match with one of the supported data types, it will be casted to a String
     * @param values List of date values to insert
     * @param cellRange Cell range as string in the format like A1:D1 or X10:X22
     * @throws FormatException Thrown if the passed address is malformed
     * @throws UndefinedStyleException Thrown if the default style was misconfigured
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    public void addBooleanCellRange(List<Boolean> values, String cellRange)
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
     * @throws UnknownRangeException Thrown if the next cell is out of range (on row or column)
     */
    private <T> void addCellRangeInternal(List<T> values, Cell.Address startAddress, Cell.Address endAddress)
    {
        List<Cell.Address> addresses = Cell.getCellRange(startAddress, endAddress);
        if (values.size() != addresses.size())
        {
            throw new UnknownRangeException("The number of passed values (" + Integer.toString(values.size()) + ") differs from the number of cells within the range (" + Integer.toString(addresses.size()) + ")");
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
     * @throws UnknownRangeException Thrown if the resolved cell address is out of range
     */
    public boolean removeCell(int columnAddress, int rowAddress)
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
     * @param address Cell address in the format A1 - XFD1048576
     * @return Returns true if the cell could be removed (existed), otherwise false (did not exist)
     * @throws UnknownRangeException Thrown if the resolved cell address is out of range
     * @throws FormatException Thrown if the passed address is malformed
     */
    public boolean removeCell(String address)
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
     * @throws UnknownRangeException Thrown if the address is out of the valid range. Range is from 0 to 1048575 (1048576 rows)
     */
    public void setCurrentRowAddress(int rowAddress)
    {
        if (rowAddress >= 1048576 || rowAddress < 0)
        {
            throw new UnknownRangeException("The row number (" + Integer.toString(rowAddress) + ") is out of range. Range is from 0 to 1048575 (1048576 rows).");
        }
        this.currentRowNumber = rowAddress;
    }
    
    /**
     * Sets the current column address (column number, zero based)
     * @param columnAddress Column number (zero based)
     * @throws UnknownRangeException Thrown if the address is out of the valid range. Range is from 0 to 16383 (16384 columns)
     */
    public void setCurrentColumnAddress(int columnAddress)
    {
        if (columnAddress >= 16383 || columnAddress < 0)
        {
            throw new UnknownRangeException("The column number (" + Integer.toString(columnAddress) + ") is out of range. Range is from 0 to 16383 (16384 columns).");
        }
        this.currentColumnNumber = columnAddress;
    }   
    
    /**
     * Set the current cell address
     * @param address Cell address in the format A1 - XFD1048576
     * @throws UnknownRangeException Thrown if the address is out of the valid range. Range is for rows from 0 to 1048575 (1048576 rows) and for columns from 0 to 16383 (16384 columns)
     * @throws FormatException Thrown if the passed address is malformed
     */
    public void setCurentCellAddress(String address)
    {
        Cell.Address adr = Cell.resolveCellCoordinate(address);
        setCurentCellAddress(adr.Column, adr.Row);
    }  
    
    /**
     * Set the current cell address
     * @param columnAddress Column number (zero based)
     * @param rowAddress Row number (zero based)
     * @throws UnknownRangeException Thrown if the address is out of the valid range. Range is for rows from 0 to 1048575 (1048576 rows) and for columns from 0 to 16383 (16384 columns)
     */
    public void setCurentCellAddress(int columnAddress, int rowAddress)
    {
        setCurrentColumnAddress(columnAddress);
        setCurrentRowAddress(rowAddress);
    }    
    
    /**
     * Set the current cell address
     * @param columnAddress Column number (zero based)
     * @param width Row number (zero based)
     * @throws UnknownRangeException Thrown if the address is out of the valid range. Range is from 0 to 16383 (16384 columns)
     */
    public void SetColumnWidth(String columnAddress, float width)
    {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnWidth(columnNumber, width);
    }
    
    /**
     * Sets the width of the passed column number (zero-based)
     * @param columnNumber Column number (zero-based, from 0 to 16383)
     * @param width Width from 0 to 255.0
     * @throws UnknownRangeException Thrown if the address is out of the valid range. Range is from 0 to 16383 (16384 columns)
     */
    public void setColumnWidth(int columnNumber, float width)
    {
        if (columnNumber >= 16384 || columnNumber < 0)
        {
            throw new UnknownRangeException("The column number (" + Integer.toString(columnNumber) + ") is out of range. Range is from 0 to 16383 (16384 columns).");
        }
        if (width < 0 || width > 255)
        {
            throw new UnknownRangeException("The column width (" + Float.toString(width) + ") is out of range. Range is from 0 to 255 (chars).");
        }
        if (this.columns.containsKey(columnNumber))
        {
            this.columns.get(columnNumber).setWidth(width);
        }
        else
        {
            Column c = new Column(columnNumber);
            c.setWidth(width);
            this.columns.put(columnNumber, c);
        }
    }  
   
    /**
     * Sets the height of the passed row number (zero-based)
     * @param rowNumber Row number (zero-based, 0 to 1048575)
     * @param height Height from 0 to 409.5
     * @throws UnknownRangeException Thrown if the address is out of the valid range. Range is from 0 to 1048575 (1048576 rows)
     */
    public void setRowHeight(int rowNumber, float height)
   {
       if (rowNumber >= 1048576 || rowNumber < 0)
       {
           throw new UnknownRangeException("The row number (" + Integer.toString(rowNumber) + ") is out of range. Range is from 0 to 1048575 (1048576 rows).");
       }
       if (height < 0 || height > 409.5)
       {
           throw new UnknownRangeException("The row height (" + Float.toString(height) + ") is out of range. Range is from 0 to 409.5 (equals 546px).");
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
    
    /**
     * Merges the defined cell range
     * @param cellRange Range to merge
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     */
    public String mergeCells(Cell.Range cellRange)
    {
        return mergeCells(cellRange.StartAddress, cellRange.EndAddress);
    }

    /**
     * Merges the defined cell range
     * @param cellRange Range to merge (e.g. 'A1:B12')
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     * @throws picoxlsx4j.exception.FormatException Thrown if the passed address is malformed
     */
    public String mergeCells(String cellRange)
    {
        Cell.Range range = Cell.resolveCellRange(cellRange);
        return mergeCells(range.StartAddress, range.EndAddress);
    }    
    
    /**
     * Merges the defined cell range
     * @param startAddress Start address of the merged cell range
     * @param endAddress End address of the merged cell range
     * @return Returns the validated range of the merged cells (e.g. 'A1:B12')
     */
    public String mergeCells(Cell.Address startAddress, Cell.Address endAddress)
    {
        List<Cell.Address> addresses = Cell.getCellRange(startAddress, endAddress);
        String key = startAddress.toString() + ":" + endAddress.toString();
        Cell.Range value = new Cell.Range(startAddress, endAddress);
        if (this.mergedCells.containsKey(key) == false)
        {
            this.mergedCells.put(key, value);
        }
        return key;
    }    
   
    /**
     * Removes the defined merged cell range
     * @param range Cell range to remove the merging
     * @throws UnknownRangeException Thrown if the passed cell range was not merged earlier
     * @throws FormatException Thrown if the passed address is malformed
     */
    public void removeMergedCells(String range)
    {
        range = range.toUpperCase();
        if (this.mergedCells.containsKey(range) == false)
        {
            throw new UnknownRangeException("The cell range " + range + " was not found in the list of merged cell ranges");
        }
        else
        {
            List<Cell.Address> addresses = Cell.getCellRange(range);
            Cell cell;
            //foreach(Cell.Address address in addresses)
            for(int i = 0; i < addresses.size(); i++)
            {
                if (this.cells.containsKey(addresses.toString()))
                {
                    cell = this.cells.get(this.cells.get(i).toString()); 
                    cell.setFieldType(Cell.CellType.DEFAULT); // resets the type
                    if (cell.getValue() == null)
                    {
                        cell.setValue("");
                    }
                }
            }
            this.mergedCells.remove(range);
        }
    }
    
    /**
     * Method to add allowed actions if the worksheet is protected. If one or more values are added, UseSheetProtection will be set to true
     * @param typeOfProtection Allowed action on the worksheet or cells
     */
    public void addAllowedActionOnSheetProtection(SheetProtectionValue typeOfProtection)
    {
        if (this.sheetProtectionValues.contains(typeOfProtection) == false)
        {
            if (typeOfProtection == SheetProtectionValue.selectLockedCells && this.sheetProtectionValues.contains(SheetProtectionValue.selectUnlockedCells) == false)
            {
                this.sheetProtectionValues.add(SheetProtectionValue.selectUnlockedCells);
            }
            this.sheetProtectionValues.add(typeOfProtection);
            this.setUseSheetProtection(true);
        }
    }
    
    /**
     * Sets or removes the password for worksheet protection. If set, UseSheetProtection will be also set to true
     * @param password Password (UTF-8) to protect the worksheet. If the password is null or empty, no password will be used
     */
    public void setSheetProtectionPassword(String password)
    {
        if (Helper.isNullOrEmpty(password) == true)
        {
            this.sheetProtectionPassword = null;
        }
        else
        {
            this.sheetProtectionPassword = password;
            this.useSheetProtection = true;
        }
    }    
    
    /**
     * Sets the defined row as hidden
     * @param rowNumber Row number to hide on the worksheet
     * @exception OutOfRangeException Thrown if the passed column number was out of range
     */
    public void addHiddenRow(int rowNumber)
    {
        setRowHiddenState(rowNumber, true);
    }
    
    /**
     * Sets a previously defined, hidden row as visible again
     * @param rowNumber Row number to hide on the worksheet
     * @exception OutOfRangeException Thrown if the passed column number was out of range
     */
    public void removeHiddenRow(int rowNumber)
    {
        setRowHiddenState(rowNumber, false);
    }
    
    /**
     * Sets the defined row as hidden or visible
     * @param rowNumber Row number to hide on the worksheet
     * @param state If true, the row will be hidden, otherwise visible
     * @exception OutOfRangeException Thrown if the passed column number was out of range
     */
    private void setRowHiddenState(int rowNumber, boolean state)
    {
        if (rowNumber >= 1048576 || rowNumber < 0)
        {
            throw new OutOfRangeException("The row number (" + Integer.toString(rowNumber) + ") is out of range. Range is from 0 to 1048575 (1048576 rows).");
        }
        if (this.hiddenRows.containsKey(rowNumber))
        {
            if (state == true)
            {
                this.hiddenRows.put(rowNumber, state);
            }
            else
            {
                this.hiddenRows.remove(rowNumber);
            }
        }
        else if (state == true)
        {
            this.hiddenRows.put(rowNumber, state);
        }
    }
    
    /**
     * Sets the defined column as hidden
     * @param columnNumber Column number to hide on the worksheet
     * @exception OutOfRangeException Thrown if the passed row number was out of range
     */
    public void addHiddenColumn(int columnNumber)
    {
        setColumnHiddenState(columnNumber, true);
    }
    
    /**
     * Sets the defined column as hidden
     * @param columnAddress Column address to hide on the worksheet
     * @exception OutOfRangeException Thrown if the passed row number was out of range
     */
    public void addHiddenColumn(String columnAddress)
    {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnHiddenState(columnNumber, true);
    }
    
    /**
     * Sets a previously defined, hidden column as visible again
     * @param columnNumber Column number to make visible again
     * @exception OutOfRangeException Thrown if the passed row number was out of range
     */
    public void removeHiddenColumn(int columnNumber)
    {
        setColumnHiddenState(columnNumber, false);
    }
    
    /**
     * Sets a previously defined, hidden column as visible again
     * @param columnAddress Column address to make visible again
     * @exception OutOfRangeException Thrown if the passed row number was out of range
     */
    public void removeHiddenColumn(String columnAddress)
    {
        int columnNumber = Cell.resolveColumn(columnAddress);
        setColumnHiddenState(columnNumber, false);
    }
    
    /**
     * Sets the defined column as hidden or visible
     * @param columnNumber Column number to hide on the worksheet
     * @param state If true, the column will be hidden, otherwise be visible
     * @exception OutOfRangeException Thrown if the passed row number was out of range
     */
    private void setColumnHiddenState(int columnNumber, boolean state)
    {
        if (columnNumber >= 16384 || columnNumber < 0)
        {
            throw new OutOfRangeException("The column number (" + Integer.toString(columnNumber) + ") is out of range. Range is from 0 to 16383 (16384 columns).");
        }
        if (this.columns.containsKey(columnNumber) && state == true)
        {
            this.columns.get(columnNumber).setHidden(state);
        }
        else if (state == true)
        {
            Column c = new Column(columnNumber);
            c.setHidden(state);
            this.columns.put(columnNumber, c);
        }
    }
    
    /**
     * Sets the column auto filter within the defined column range
     * @param startColumn Column number with the first appearance of a auto filter drop down
     * @param endColumn Column number with the last appearance of a auto filter drop down
     * @exception OutOfRangeException Thrown if one of the passed column numbers are out of range
     */
    public void setAutoFilter(int startColumn, int endColumn)
    {
        if (startColumn >= 16384 || startColumn < 0)
        {
            throw new OutOfRangeException("The start column number (" + Integer.toString(startColumn) + ") is out of range. Range is from 0 to 16383 (16384 columns).");
        }
        if (endColumn >= 16384 || endColumn < 0)
        {
            throw new OutOfRangeException("The end column number (" + Integer.toString(endColumn) + ") is out of range. Range is from 0 to 16383 (16384 columns).");
        }
        String start = Cell.resolveCellAddress(startColumn, 0);
        String end = Cell.resolveCellAddress(endColumn, 0);
        if (endColumn < startColumn)
        {
            setAutoFilter(end + ":" + start);
        }
        else
        {
            setAutoFilter(start + ":" + end);
        }
    }
    
    /**
     * Sets the column auto filter within the defined column range
     * @param range Range to apply auto filter on. The range could be 'A1:C10' for instance. The end row will be recalculated automatically when saving the file
     * @exception OutOfRangeException Thrown if the passed range out of range
     * @exception FormatException Thrown if the passed range is malformed
     */
    public void setAutoFilter(String range)
    {
        this.autoFilterRange = Cell.resolveCellRange(range);
        recalculateAutoFilter();
        recalculateColumns();
    } 
    
    /**
     * Removes auto filters from the worksheet
     */
    public void removeAutoFilter()
    {
        this.autoFilterRange = null;
    }    

    /**
     * Method to recalculate the auto filter (columns) of this worksheet. This is an internal method. There is no need to use it. It must be public to require access from the LowLevel class
     */
    public void recalculateAutoFilter()
    {
       if (this.autoFilterRange == null) { return; }
       int start = this.autoFilterRange.StartAddress.Column;
       int end = this.autoFilterRange.EndAddress.Column;
       int endRow = 0;
       for(Map.Entry<String, Cell> item  : this.getCells().entrySet())
       {
           if (item.getValue().getColumnAddress() < start || item.getValue().getColumnAddress() > end) { continue; }
           if (item.getValue().getRowAddress() > endRow) {endRow = item.getValue().getRowAddress();}
       }
       Column c;
       for(int i = start; i <= end; i++)
       {
           if (this.columns.containsKey(i) == false)
           {
               c = new Column(i);
               c.setAutoFilter(true);
               this.columns.put(i, c);
           }
           else
           {
               this.getColumns().get(i).setAutoFilter(true);
           }
       }
       Cell.Range temp = new Cell.Range(new Cell.Address(start, 0), new Cell.Address(end, endRow));
       this.autoFilterRange = temp;       
    }
    
    /**
     * Method to recalculate the collection of columns of this worksheet. This is an internal method. There is no need to use it. It must be public to require access from the LowLevel class
     */
    public void recalculateColumns()
    {
       ArrayList<Integer> columnsToDelete = new ArrayList<>();
       for(Map.Entry<Integer, Column> col  : this.getColumns().entrySet())
       {
           if (col.getValue().hasAutoFilter() == false && col.getValue().isHidden() == false && col.getValue().getWidth() != Worksheet.DEFAULT_COLUMN_WIDTH)
           {
               columnsToDelete.add(col.getKey());
           }
       }
        for(Iterator<Integer> index = columnsToDelete.iterator(); index.hasNext(); )
        {
            this.columns.remove(index.next());
        }       
    }
    
    /**
     * Sets the selected cells on this worksheet
     * @param range Cell range to select
     */
    public void setSelectedCells(Cell.Range range)
    {
        this.selectedCells = range;
    }
    
    /**
     * Sets the selected cells on this worksheet
     * @param startAddress Start address of the range
     * @param endAddress End address of the range
     */
    public void setSelectedCells(Cell.Address startAddress, Cell.Address endAddress)
    {
       this.selectedCells = new Cell.Range(startAddress, endAddress); 
    }
    
    /**
     * Sets the selected cells on this worksheet
     * @param range Cell range to select
     * @exception OutOfRangeException Thrown if the passed range out of range
     * @exception FormatException Thrown if the passed range is malformed
     */
    public void setSelectedCells(String range)
    {
        this.selectedCells = Cell.resolveCellRange(range);
    }  
    
    /**
     * Removes the cell selection of this worksheet
     */
    public void removeSelectedCells()
    {
        this.selectedCells = null;
    }
    
}
