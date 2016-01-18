/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j;

/**
 * Class representing a column of a worksheet
 * @author Raphael Stoeckli
 */
public class Column {
    
    private int number;
    private String columnAddress;
    private float width;
    private boolean hidden;
    private boolean autoFilter;

    /**
     * Gets the column number
     * @return Column number (0 to 16383)
     */
    public int getNumber() {
        return number;
    }

    /**
     * Sets the column number
     * @param number Column number (0 to 16383)
     */
    public void setNumber(int number) {
        this.columnAddress = Cell.resolveColumnAddress(number);
        this.number = number;
    }

    /**
     * Gets the column address
     * @return Column address (A to XFD)
     */
    public String getColumnAddress() {
        return columnAddress;
    }

    /**
     * Sets the column address
     * @param columnAddress Column address (A to XFD)
     */
    public void setColumnAddress(String columnAddress) {
        this.number = Cell.resolveColumn(columnAddress);
        this.columnAddress = columnAddress;
    }

    /**
     * Gets the width of the column
     * @return Width of the column
     */
    public float getWidth() {
        return width;
    }

    /**
     * Sets the width of the column
     * @param width Width of the column
     */
    public void setWidth(float width) {
        this.width = width;
    }

    /**
     * Gets whether the column is hidden or visible
     * @return If true, the column is hidden, otherwise visible
     */
    public boolean isHidden() {
        return hidden;
    }

    /**
     * Sets whether the column is hidden or visible
     * @param isHidden If true, the column is hidden, otherwise visible
     */
    public void setHidden(boolean isHidden) {
        this.hidden = isHidden;
    }

    /**
     * Gets whether auto filter is enabled on the column
     * @return If true, the column has auto filter applied, otherwise not
     */
    public boolean hasAutoFilter() {
        return autoFilter;
    }

    /**
     * Sets whether auto filter is enabled on the column
     * @param hasAutoFilter If true, the column has auto filter applied, otherwise not
     */
    public void setAutoFilter(boolean hasAutoFilter) {
        this.autoFilter = hasAutoFilter;
    }

    /**
     * Default constructor
     */
    public Column() {
        this.width = Worksheet.DEFAULT_COLUMN_WIDTH;
    }

    
    /**
     * Constructor with column address
     * @param columnAddress Column address (A to XFD)
     */
    public Column(String columnAddress) {
        this.setColumnAddress(columnAddress);
        this.width = Worksheet.DEFAULT_COLUMN_WIDTH;
    }

    /**
     * Constructor with column number
     * @param number Column number (zero-based, 0 to 16383)
     */
    public Column(int number) {
        this.setNumber(number);
        this.width = Worksheet.DEFAULT_COLUMN_WIDTH;
    }
    
}
