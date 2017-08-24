/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2017
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j;

import picoxlsx4j.style.Style;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import picoxlsx4j.exception.FormatException;
import picoxlsx4j.exception.UnknownRangeException;
import picoxlsx4j.exception.UndefinedStyleException;
import java.util.regex.*;
/*
import picoxlsx4j.Address;
import picoxlsx4j.Helper;
import picoxlsx4j.Worksheet;
*/
import picoxlsx4j.exception.OutOfRangeException;

/**
 * Class representing a cell of a worksheet
 * @author Raphael Stoeckli
 */
public class Cell implements Comparable<Cell>{
    
// ### E N U M S ###
    /**
     * Enum defines the basic data types of a cell
     */
    public enum CellType
    {
        /**
         * Type for single characters and strings
         */
        STRING,
        /**
         * Type for all numeric types (long, integer and float and double)
         */
        NUMBER,
        /**
         * Type for dates and times (Note: Dates before 1900-01-01 are not allowed)
         */
        DATE,
        /**
         * Type for boolean
         */
        BOOL,
        /**
         * Type for Formulas (The cell will be handled differently)
         */
        FORMULA,
        /**
         * Type for empty cells. This type is only used for merged cells (all cells except the first of the cell range)
         */
        EMPTY,
        /**
         * Default Type, not specified
         */
        DEFAULT
    }    
    
// ### P R I V A T E  F I E L D S ###
    
    private Style cellStyle;
    private int columnAddress;
    private CellType fieldType;
    private int rowAddress;
    private Object value;
    private Worksheet worksheetReference;
    
// ### G E T T E R S  &  S E T T E R S ###

 /**
     * Gets the combined cell Address as string in the format A1 - XFD1048576
     * @return Cell address
     */
    public String getCellAddress()
    {
        return Cell.resolveCellAddress(this.columnAddress, this.rowAddress);
    }
    /**
     * Sets the combined cell Address as string in the format A1 - XFD1048576
     * @param address Cell address
     * @throws UnknownRangeException Thrown in case of a illegal address
     */
    public void setCellAddress(String address)
    {
        Address temp = Cell.resolveCellCoordinate(address);
        this.columnAddress = temp.Column;
        this.rowAddress = temp.Row;
    }
    /**
     * Gets the combined cell address as class
     * @return Cell address
     */
    public Address getCellAddress2()
    {
        return new Address(this.columnAddress, this.rowAddress);
    }
    /**
     * Sets the combined cell address as class
     * @param address Cell address
     */
    public void setCellAddress2(Address address)
    {
        this.setColumnAddress(address.Column);
        this.setRowAddress(address.Row);
    }

    /**
     * Gets the assigned style of the cell
     * @return Assigned style
     */
    public Style getCellStyle() {
        return cellStyle;
    }


    /**
     * Gets the number of the column (zero-based)
     * @return Column number (zero-based)
     */    
    public int getColumnAddress() {
        return columnAddress;
    }

    /**
     * Sets the number of the column (zero-based)
     * @param columnAddress Column number (zero-based)
     */
    public void setColumnAddress(int columnAddress) {
        if (columnAddress < Worksheet.MIN_COLUMN_ADDRESS || columnAddress > Worksheet.MAX_COLUMN_ADDRESS)
        {
            throw new OutOfRangeException("The passed number (" + Integer.toString(columnAddress) + ")is out of range. Range is from " + Integer.toString(Worksheet.MIN_COLUMN_ADDRESS) + " to " + Integer.toString(Worksheet.MAX_COLUMN_ADDRESS) + " (" + (Integer.toString(Worksheet.MAX_COLUMN_ADDRESS + 1)) + " rows).");
        }        
        this.columnAddress = columnAddress;
    }

    /**
     * Gets the type of the cell
     * @return Type of the cell
     */
    public CellType getFieldType() {
        return fieldType;
    }

    /**
     * Sets the type of the cell
     * @param fieldType Type of the cell
     */
    public void setFieldType(CellType fieldType) {
        this.fieldType = fieldType;
    }
    /**
     * Gets the number of the row (zero-based)
     * @return Row number (zero-based)
     */
    public int getRowAddress() {
        return rowAddress;
    }
    /**
     * Sets the number of the row (zero-based)
     * @param rowAddress Row number (zero-based)
     */
    public void setRowAddress(int rowAddress) {
        if (rowAddress < Worksheet.MIN_ROW_ADDRESS || rowAddress > Worksheet.MAX_ROW_ADDRESS)
        {
            throw new OutOfRangeException("The passed number (" + Integer.toString(rowAddress) + ")is out of range. Range is from " + Integer.toString(Worksheet.MIN_ROW_ADDRESS) + " to " + Integer.toString(Worksheet.MAX_ROW_ADDRESS) + " (" + (Integer.toString(Worksheet.MAX_ROW_ADDRESS + 1)) + " rows).");
        }
        this.rowAddress = rowAddress;
    }
    /**
     * Gets the value of the cell (generic object type)
     * @return Value of the cell
     */
    public Object getValue() {
        return value;
    }
    /**
     * Sets the value of the cell (generic object type)
     * @param value Value of the cell
     */
    public void setValue(Object value) {
        this.value = value;
    } 
    
    /**
     * Gets or sets the parent worksheet reference
     * @return Worksheet reference
     */
    public Worksheet getWorksheetReference()
    {
        return this.worksheetReference;
    }
    
    /**
     * Sets the parent worksheet reference
     * @param reference Worksheet reference
     */
    public void setWorksheetReference(Worksheet reference)
    {
        this.worksheetReference = reference;
    }
    
    
// ### C O N S T R U C T O R S ###
    
    /**
     * Default constructor
     */
    public Cell()
    {
        this.worksheetReference = null;
    }
    /**
     * Constructor with value and cell type
     * @param value Value of the cell
     * @param type Type of the cell
     */
    public Cell(Object value, CellType type)
    {
        this.fieldType = type;
        this.value = value;
        resolveCellType();
    }
    /**
     * Constructor with value, cell type, row address and column address
     * @param value Value of the cell
     * @param type Type of the cell
     * @param column Column address of the cell (zero-based)
     * @param row Row address of the cell (zero-based)
     * @param reference Worksheet reference
     */
    public Cell(Object value, CellType type, int column, int row, Worksheet reference)
    {
        this.fieldType = type;
        this.value = value;
        this.columnAddress = column;
        this.rowAddress = row;
        this.worksheetReference = reference;
        if (type == CellType.DEFAULT)
        {
            resolveCellType();
        }
    }
    
// ### M E T H O D S ###
    
    /**
     * Implemented compareTo method
     * @param o Object to compare
     * @return 0 if values are the same, -1 if this object is smaller, 1 if it is bigger
     */
    @Override
    public int compareTo(Cell o) {
        if (this.rowAddress == o.rowAddress)
        {
            return Integer.compare(this.columnAddress, o.getColumnAddress());
        }
        else
        {
            return Integer.compare(this.rowAddress, o.getRowAddress());
        }
    }
   
    /**
     * Removes the assigned style from the cell
     * @throws UndefinedStyleException Thrown if the workbook to remove was not found in the style sheet collection
     */
    public void removeStyle()
    {
        if (this.worksheetReference == null)
        {
            throw new UndefinedStyleException("No worksheet reference was defined while trying to remove a style from a cell");
        }
        if (this.worksheetReference.getWorkbookReference() == null)
        {
            throw new UndefinedStyleException("No workbook reference was defined on the worksheet while trying to remove a style from a cell");
        }
        if (this.cellStyle != null)
        {
            String styleName = this.cellStyle.getName();
            this.cellStyle = null;
            this.worksheetReference.getWorkbookReference().removeStyle(styleName, true);
        }
    }
    
     /**
      * Method resets the Cell type and tries to find the actual type. This is used if a Cell was created with the CellType DEFAULT. CellTypes FORMULA and EMPTY will skip this method
      */
    public void resolveCellType()
    {
        if(this.value == null)
        {
            this.setFieldType(CellType.EMPTY);
            value = "";
            return;
        }        
        if (this.fieldType == CellType.FORMULA || this.fieldType == CellType.EMPTY) {return;}
        if (value instanceof Integer) { this.fieldType = CellType.NUMBER; }
        else if (value instanceof Long) { this.fieldType = CellType.NUMBER; }
        else if (value instanceof Float) { this.fieldType = CellType.NUMBER; }
        else if (value instanceof Double) { this.fieldType = CellType.NUMBER; }
        else if (value instanceof Boolean) { this.fieldType = CellType.BOOL; }
        else if (value instanceof Date) { this.fieldType = CellType.DATE; }
        else { this.fieldType = CellType.STRING; } // Default
    }
    /**
     * Sets the lock state of the cell
     * @param isLocked If true, the cell will be locked if the worksheet is protected
     * @param isHidden If true, the value of the cell will be invisible if the worksheet is protected
     */
    public void setCellLockedState(boolean isLocked, boolean isHidden)
    {
        Style lockStyle;
        if (this.cellStyle == null)
        {
            lockStyle = new Style();
        }
        else
        {
            lockStyle = this.cellStyle.copy();
        }
        lockStyle.getCurrentCellXf().setLocked(isLocked);
        lockStyle.getCurrentCellXf().setHidden(isHidden);
        try
        {
            this.setStyle(lockStyle);
        }
        catch(Exception e)
        {
            // Should never happen
        }
    }
    
    /**
     * Sets the style of the cell
     * @param style style to assign
     * @return If the passed style already exists in the workbook, the existing one will be returned, otherwise the passed one
     * @throws UndefinedStyleException Thrown if the style is not referenced in the workbook
     */
    public Style setStyle(Style style)
    {
       if (this.worksheetReference == null)
       {
           throw new UndefinedStyleException("No worksheet reference was defined while trying to set a style to a cell");
       }
       if (this.worksheetReference.getWorkbookReference() == null)
       {
           throw new UndefinedStyleException("No workbook reference was defined on the worksheet while trying to set a style to a cell");
       }
       if (style == null)
       {
           throw new UndefinedStyleException("No style to assign was defined");
       }
       Style s = this.worksheetReference.getWorkbookReference().addStyle(style, true);
       this.cellStyle = s;
       return s;
    }
    
// ### S T A T I C   M E T H O D S ###
    
    /**
     * Get a list of cell addresses from a cell range
     * @param startColumn Start column (zero based)
     * @param startRow Start row (zero based)
     * @param endColumn End column (zero based)
     * @param endRow End row (zero based)
     * @return List of cell addresses
     */
    public static List<Address> GetCellRange(int startColumn, int startRow, int endColumn, int endRow)
    {
        Address start = new Address(startColumn, startRow);
        Address end = new Address(endColumn, endRow);
        return getCellRange(start, end);       
    }
    
    /**
     * Converts a List of supported objects into a list of cells
     * @param <T> Generic data type
     * @param list List of generic objects
     * @return List of cells
     */
    public  static <T> List<Cell> convertArray(List<T> list)
    {
        List<Cell> output = new ArrayList<>();
        Cell c = null;
        Object o = null;
        for (int i = 0; i < list.size(); i++)
        {
            o = list.get(i);
            if (o instanceof Integer)
            {
                c = new Cell(o, CellType.NUMBER);
            }
            else if (o instanceof Long)
            {
                c = new Cell(o, CellType.NUMBER);
            }
            else if (o instanceof Float)
            {
                c = new Cell(o, CellType.NUMBER);
            }
            else if (o instanceof Double)
            {
                c = new Cell(o, CellType.NUMBER);
            }
            else if (o instanceof Boolean)
            {
                c = new Cell(o, CellType.BOOL);
            }
            else if (o instanceof Date)
            {
                c = new Cell(o, CellType.DATE);
            }
            else if (o instanceof String)
            {
                c = new Cell(o, CellType.STRING);
            }
            else
            {
                c = new Cell(o, CellType.DEFAULT);
                //throw new UnsupportedDataTypeException("The data type '" + t.toString() + "' is not supported");
            }
            output.add(c);
        }
        return output;
    }
    
    /**
     * Gets a list of cell addresses from a cell range (format A1:B3 or AAD556:AAD1000)
     * @param range Range to process
     * @return List of cell addresses
     * @throws FormatException Thrown if the passed address range is malformed
     */
    public static List<Address> getCellRange(String range)
    {
       Range range2 = resolveCellRange(range);
       return getCellRange(range2.StartAddress, range2.EndAddress);
    }
    
    /**
     * Get a list of cell addresses from a cell range
     * @param startAddress Start address as string in the format A1 - XFD1048576
     * @param endAddress End address as string in the format A1 - XFD1048576
     * @return List of cell addresses
     * @throws FormatException Thrown if one of the passed addresses contains malformed information
     * @throws UnknownRangeException Thrown if one of the passed addresses is out of range
     */
    public static List<Address> getCellRange(String startAddress, String endAddress)
    {
        Address start = resolveCellCoordinate(startAddress);
        Address end = resolveCellCoordinate(endAddress);
        return getCellRange(start, end);
    }    
    
    /**
     * Get a list of cell addresses from a cell range
     * @param startAddress Start address
     * @param endAddress End address
     * @return List of cell addresses
     */
    public static List<Address> getCellRange(Address startAddress, Address endAddress)
    {
            int startColumn, endColumn, startRow, endRow;
            if (startAddress.Column < endAddress.Column)
            {
                startColumn = startAddress.Column;
                endColumn = endAddress.Column;
            }
            else
            {
                startColumn = endAddress.Column;
                endColumn = startAddress.Column;
            }
            if (startAddress.Row < endAddress.Row)
            {
                startRow = startAddress.Row;
                endRow = endAddress.Row;
            }
            else
            {
                startRow = endAddress.Row;
                endRow = startAddress.Row;
            }
            List<Address> output = new ArrayList<>();
            for (int i = startRow; i <= endRow; i++)
            {
                for (int j = startColumn; j <= endColumn; j++)
                {
                    output.add(new Address(j, i));
                }
            }
            return output;
    }
    
    /**
     * Gets the address of a cell by the column and row number (zero based)
     * @param column Column address of the cell (zero-based)
     * @param row Row address of the cell (zero-based)
     * @return Cell Address as string in the format A1 - XFD1048576
     * @throws UnknownRangeException Thrown if the start or end address was out of range
     */
    public static String resolveCellAddress(int column, int row)
    {
            if (row > Worksheet.MAX_ROW_ADDRESS || row < Worksheet.MIN_ROW_ADDRESS)
            {
                throw new UnknownRangeException("The row number (" + Integer.toString(row) + ") is out of range. Range is from " + Integer.toString(Worksheet.MIN_ROW_ADDRESS) + " to " + Integer.toString(Worksheet.MAX_ROW_ADDRESS) + " (" + (Integer.toString(Worksheet.MIN_ROW_ADDRESS) + 1) + " rows).");
            }
            return resolveColumnAddress(column) + Integer.toString(row + 1);    
    }
    
    /**
     * Gets the column and row number (zero based) of a cell by the address
     * @param address Address as string in the format A1 - XFD1048576
     * @return Address object of the passed string
     * @throws FormatException Thrown if the passed address was malformed
     * @throws UnknownRangeException Thrown if the resolved address is out of range
     */
    public static Address resolveCellCoordinate(String address)
    {
        int row, column;
        if (Helper.isNullOrEmpty(address))
        {
            throw new FormatException("The cell address is null or empty and could not be resolved");
        }
        address = address.toUpperCase();
        Pattern pattern = Pattern.compile("([A-Z]{1,3})([0-9]{1,7})");
        Matcher mx = pattern.matcher(address);
        if (mx.groupCount() != 2)
        {
            throw new FormatException("The format of the cell address (" + address + ") is malformed");
        }
        mx.find();
        int digits = Integer.parseInt(mx.group(2));
        column = resolveColumn(mx.group(1));
        row = digits - 1;
        
        if (row > Worksheet.MAX_ROW_ADDRESS || row < Worksheet.MIN_ROW_ADDRESS)
        {
            throw new UnknownRangeException("The row number (" + Integer.toString(row) + ") is out of range. Range is from " + Integer.toString(Worksheet.MIN_ROW_ADDRESS) + " to " + Integer.toString(Worksheet.MAX_ROW_ADDRESS) + " (" + Integer.toString((Worksheet.MAX_ROW_ADDRESS + 1)) + " rows).");
        }     
        if (column > Worksheet.MAX_COLUMN_ADDRESS || column < Worksheet.MIN_COLUMN_ADDRESS)
        {
            throw new UnknownRangeException("The column number (" + Integer.toString(column) + ") is out of range. Range is from " + Integer.toString(Worksheet.MIN_COLUMN_ADDRESS) + " to " + Integer.toString(Worksheet.MAX_COLUMN_ADDRESS) + " (" + Integer.toString((Worksheet.MAX_COLUMN_ADDRESS + 1)) + " columns).");
        }
        
        Address output = new Address(column, row);
        return output;
    } 
    /**
     * Resolves a cell range from the format like A1:B3 or AAD556:AAD1000
     * @param range Range to process
     * @return Range object of the passed string range
     * @throws FormatException Thrown if the passed range is malformed
     */
    public static Range resolveCellRange(String range)
    {
        if (Helper.isNullOrEmpty(range))
        {
            throw new FormatException("The cell range is null or empty and could not be resolved");
        }
        String[] split = range.split(":");
        if (split.length != 2)
        {
            throw new FormatException("The cell range (" + range + ") is malformed and could not be resolved");
        }
        try
        {
            Address start = resolveCellCoordinate(split[0]);
            Address end = resolveCellCoordinate(split[1]);
            Range output = new Range(start, end);
            return output;
        }
        catch(Exception e)
        {
            throw new FormatException("The start address or end address could not be resolved. See inner exception", e);
        }
    }
   
    /**
     * Gets the column number from the column address (A - XFD)
     * @param columnAddress Column address (A - XFD)
     * @return Column number (zero-based)
     * @throws UnknownRangeException Thrown if the column is out of range
     */
    public static int resolveColumn(String columnAddress)
    {
        int temp;
        int result = 0;
        int multiplicator = 1;
        
        for (int i = columnAddress.length() - 1; i >= 0; i--)
        {
            temp = (int)columnAddress.charAt(i);
            temp = temp - 64;
            result = result + (temp * multiplicator);
            multiplicator = multiplicator * 26;
        }
        if (result - 1 > Worksheet.MAX_COLUMN_ADDRESS || result - 1 < Worksheet.MIN_COLUMN_ADDRESS)
        {
            throw new UnknownRangeException("The column number (" + Integer.toString(result - 1) + ") is out of range. Range is from " + Integer.toString(Worksheet.MIN_COLUMN_ADDRESS) + " to " + Integer.toString(Worksheet.MAX_COLUMN_ADDRESS) + " (" + Integer.toString((Worksheet.MAX_COLUMN_ADDRESS + 1)) + " columns).");
        }        
        return result - 1;
    }
    
    /**
     * Gets the column address (A - XFD)
     * @param columnNumber Column number (zero-based)
     * @return Column address (A - XFD)
     * @throws UnknownRangeException Thrown if the passed column number is out of range
     */
    public static String resolveColumnAddress(int columnNumber)
    {
        if (columnNumber > Worksheet.MAX_COLUMN_ADDRESS || columnNumber < Worksheet.MIN_COLUMN_ADDRESS)
        {
            throw new UnknownRangeException("The column number (" + Integer.toString(columnNumber) + ") is out of range. Range is from " + Integer.toString(Worksheet.MIN_COLUMN_ADDRESS) + " to " + Integer.toString(Worksheet.MAX_COLUMN_ADDRESS) + " (" + Integer.toString((Worksheet.MAX_COLUMN_ADDRESS + 1)) + " columns).");
        }
        // A - XFD
        int j = 0;
        int k = 0;
        int l = 0;
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i <= columnNumber; i++)
        {
            if (j > 25)
            {
                k++;
                j = 0;
            }
            if (k > 25)
            {
                l++;
                k = 0;
            }
            j++;
        }
        if (l > 0) { sb.append((char)(l + 64)); }
        if (k > 0) { sb.append((char)(k + 64)); }
        sb.append((char)(j + 64));
        return sb.toString();
    }
}
