/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j;

import picoxlsx4j.style.Style;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import picoxlsx4j.exception.FormatException;
import picoxlsx4j.exception.IOException;
import picoxlsx4j.exception.UndefinedStyleException;
import picoxlsx4j.exception.UnknownRangeException;
import picoxlsx4j.exception.UnknownWorksheetException;
import picoxlsx4j.exception.WorksheetNameAlreadxExistsException;
import picoxlsx4j.lowLevel.LowLevel;
import picoxlsx4j.style.BasicStyles;
import picoxlsx4j.style.Border;
import picoxlsx4j.style.CellXf;
import picoxlsx4j.style.Fill;
import picoxlsx4j.style.Font;
import picoxlsx4j.style.NumberFormat;
import picoxlsx4j.style.StyleCollection;

/**
 * Class representing a workbook
 * @author Raphael Stoeckli
 */
public class Workbook {
    
    private String filename;
    private List<Worksheet> worksheets;
    private Worksheet currentWorksheet;
    private List<Style> styles;
    private Metadata workbookMetadata;
    private String workbookProtectionPassword;
    private boolean lockWindowsIfProtected;
    private boolean lockStructureIfProtected;
    private boolean useWorkbookProtection;

    /**
     * Gets the current worksheet
     * @return Current worksheet reference
     */
    public Worksheet getCurrentWorksheet() {
        return currentWorksheet;
    }
    
    /**
     * Gets the list of worksheets in the workbook
     * @return List of worksheet objects
     */
    public List<Worksheet> getWorksheets() {
        return worksheets;
    } 
    
    /**
     * Gets the filename of the workbook
     * @return Filename of the workbook
     */
    public String getFilename() {
        return filename;
    }

    /**
     * Sets the filename of the workbook
     * @param filename Filename of the workbook
     */
    public void setFilename(String filename) {
        this.filename = filename;
    }

    /**
     * Gets the list of Styles of the workbook
     * @return List of Style objects (style sheet)
     */
    public List<Style> getStyles() {
        return styles;
    }
    
    /**
     * Gets the meta data object of the workbook
     * @return Meta data object
     */
    public Metadata getWorkbookMetadata() {
        return workbookMetadata;
    }

    /**
     * Sets the meta data object of the workbook
     * @param workbookMetadata Meta data object
     */
    public void setWorkbookMetadata(Metadata workbookMetadata) {
        this.workbookMetadata = workbookMetadata;
    }    

    /**
     * Gets whether the workbook is protected
     * @return If true, the workbook is protected otherwise not
     */
    public boolean isWorkbookProtectionUsed() {
        return useWorkbookProtection;
    }

    /**
     * Sets whether the workbook is protected
     * @param useWorkbookProtection If true, the workbook is protected otherwise not
     */
    public void setWorkbookProtection(boolean useWorkbookProtection) {
        this.useWorkbookProtection = useWorkbookProtection;
    }

    /**
     * Gets the password used for workbook protection
     * @return Password (UTF-8)
     */
    public String getWorkbookProtectionPassword() {
        return workbookProtectionPassword;
    }
    
    /**
     * Gets whether the windows are locked if workbook is protected
     * @return True if the windows are locked when the workbook is protected
     */
    public boolean isWindowsLockedIfProtected() {
        return lockWindowsIfProtected;
    }

    /**
     * Gets whether the structure are locked if workbook is protected
     * @return True if the structure is locked when the workbook is protected
     */
    public boolean isStructureLockedIfProtected() {
        return lockStructureIfProtected;
    }
    
    /**
     * Default Constructor with additional parameter to create a default worksheet
     * @param createWorksheet If true, a default worksheet will be created and set as default worksheet
     */
    public Workbook(boolean createWorksheet)
    { 
        try
        {
            this.worksheets = new ArrayList<>();
            this.styles = new ArrayList<>();
            this.styles.add(new Style("default")); // Do not remove this (Default style)
            this.styles.add(BasicStyles.DottedFill_0_125());  // Additional style to provide fill styles (compatibility?)
            this.workbookMetadata = new Metadata();
            if (createWorksheet == true)
            {
                addWorksheet("Sheet1");
            }
        }
        catch(Exception e)
        {
            // Do nothing -> Default should never throw an exception
        }
        
    }    
    
    /**
     * Constructor with filename ant the name of the first worksheet
     * @param filename Filename of the workbook
     * @param sheetName Name of the first worksheet
     * @throws WorksheetNameAlreadxExistsException thrown if the passed worksheet name already exists
     * @throws FormatException Thrown if the worksheet name contains illegal characters
     */
    public Workbook(String filename, String sheetName)
    {
        this.worksheets = new ArrayList<>();
        this.styles = new ArrayList<>();
        this.styles.add(new Style("default")); 
        this.styles.add(BasicStyles.DottedFill_0_125());
        this.workbookMetadata = new Metadata();        
        this.filename = filename;
        addWorksheet(sheetName);
    }    
    
    /**
    * Adding a new Worksheet
    * @param name Name of the new worksheet
    * @throws WorksheetNameAlreadxExistsException Thrown if the name of the worksheet already exists
    * @throws FormatException Thrown if the worksheet name contains illegal characters or is out of range (length between 1 an 31
    */
    public void addWorksheet(String name)
    {
        for (int i = 0; i < this.worksheets.size(); i++)
        {
            if (this.worksheets.get(i).getSheetName().equals(name))
            {
                throw new WorksheetNameAlreadxExistsException("The worksheet with the name '" + name + "' already exists.");
            }
        }
        int number = this.worksheets.size() + 1;
        this.currentWorksheet = new Worksheet(name, number);
        this.worksheets.add(this.currentWorksheet);
    }    

    /**
    * Sets the current worksheet
    * @param name Name of the worksheet
    * @return Returns the current worksheet
    * @throws UnknownWorksheetException Thrown if the name of the worksheet is unknown
    */   
    public Worksheet setCurrentWorksheet(String name)
    {
        boolean exists = false;
        for(int i = 0; i < this.worksheets.size(); i++)
        {
            if (this.worksheets.get(i).getSheetName().equals(name))
            {
                this.currentWorksheet = this.worksheets.get(i);
                exists = true;
                break;
            }
        }
        if (exists == false)
        {
            throw new UnknownWorksheetException("The worksheet with the name '" + name + "' does not exist.");
        }
        return this.currentWorksheet;
    }
    
    /**
    * Removes the defined worksheet
    * @param name Name of the worksheet
    * @throws UnknownWorksheetException Thrown if the name of the worksheet is unknown
    */
    public void removeWorksheet(String name)
    {
        boolean exists = false;
        boolean resetCurrent = false;
        int index = 0;
        for (int i = 0; i < this.worksheets.size(); i++)
        {
            if (this.worksheets.get(index).getSheetName().equals(name))
            {
                index = i;
                exists = true;
                break;
            }
        }
        if (exists == false)
        {
            throw new UnknownWorksheetException("The worksheet with the name '" + name + "' does not exist.");
        }
        if (this.worksheets.get(index).getSheetName().equals(this.currentWorksheet.getSheetName()) )
        {
            resetCurrent = true;
        }
        this.worksheets.remove(index);
        if (this.worksheets.size() > 0)
        {
            for (int i = 0; i < this.worksheets.size(); i++)
            {
                this.worksheets.get(index).setSheetID(i + 1);
                if (resetCurrent == true && i == 0)
                {
                    this.currentWorksheet = this.worksheets.get(i);
                }
            }
        }
        else
        {
            this.currentWorksheet = null;
        }        
    }
    
    /**
     * Sets or removes the workbook protection. If protectWindows and protectStructure are both false, the workbook will not be protected
     * @param state If true, the workbook will be protected, otherwise not
     * @param protectWindows If true, the windows will be locked if the workbook is protected
     * @param protectStructure If true, the structure will be locked if the workbook is protected
     * @param password Optional password. If null or empty, no password will be set in case of protection
     */
    public void setWorkbookProtection(boolean state, boolean protectWindows, boolean protectStructure, String password)
    {
        this.lockWindowsIfProtected = protectWindows;
        this.lockStructureIfProtected = protectStructure;
        this.workbookProtectionPassword = password;
        if (protectWindows == false && protectStructure == false)
        {
            this.useWorkbookProtection = false;
        }
        else
        {
            this.useWorkbookProtection = state;
        }
    }    
    
    /**
    * Adds a style to the style sheet of the workbook
    * @param style Style to add
    * @param distinct If true, the passed style will be replaced by an identical style if existing. Otherwise an exception will be thrown in case of a duplicate
    * @return Returns the added style. In case of an existing style, the distinct style will be returned
    * @throws UndefinedStyleException Thrown if the style could not be added to the style sheet
    */
    public Style addStyle(Style style, boolean distinct)
    {
            boolean styleExits = false;
            boolean identicalStyle = false;
            Style s;
            for (int i = 0; i < this.styles.size(); i++)
            {
                if (this.styles.get(i).getName().equals(style.getName()))
                {
                    if (this.styles.get(i).equals(style) && distinct == true)
                    {
                        identicalStyle = true;
                        s = this.styles.get(i);
                    }
                    styleExits = true;
                    break;
                }
            }
            if (styleExits == true)
            {
                if (distinct == false && identicalStyle == false)
                {
                    throw new UndefinedStyleException("The style with the name '" + style.getName() + "' already exits");
                }
                else
                {
                    s = style;
                }
            }
            else
            {
                s = style;
                this.styles.add(style);
            }
            return s;        
    }
    
    /**
    * Removes the passed style from the style sheet
    * @param style Style to remove
    * @throws UndefinedStyleException Thrown if the style is not defined in the style sheet
    */
    public void removeStyle(Style style)
    {
        removeStyle(style, false);
    }
    
    /**
    * Removes the defined style from the style sheet of the workbook
    * @param styleName Name of the style to be removed
    * @throws UndefinedStyleException Thrown if the style is not defined in the style sheet
    */
    public void removeStyle(String styleName)
    {
        removeStyle(styleName, false);
    }    

    /**
    * Removes the defined style from the style sheet of the workbook
    * @param style Style to remove
    * @param onlyIfUnused If true, the style will only be removed if not used in any cell
    * @throws UndefinedStyleException Thrown if the style is not defined in the style sheet
    */
    public void removeStyle(Style style, boolean onlyIfUnused)
    {
        if (style == null)
        {
            throw new UndefinedStyleException("The style to remove is not defined");
        }
        removeStyle(style.getName(), onlyIfUnused);
    }
    
    /**
    * Removes the defined style from the style sheet of the workbook
    * @param styleName Name of the style to be removed
    * @param onlyIfUnused If true, the style will only be removed if not used in any cell
    * @throws UndefinedStyleException Thrown if the style is not defined in the style sheet
    */
    public void removeStyle(String styleName, boolean onlyIfUnused)
    {
        if (Helper.isNullOrEmpty(styleName))
        {
            throw new UndefinedStyleException("The style to remove is not defined (no name specified)");
        }
        int index = -1;
        for(int i = 0; i < this.styles.size(); i++)
        {
            if (this.styles.get(i).getName().equals(styleName))
            {
                index = i;
                break;
            }
        }
        if (index < 0)
        {
            throw new UndefinedStyleException("The style with the name '" + styleName + "' to remove was not found in the list of styles");
        }
        else if (this.styles.get(index).getName().equals("default") || index == 0)
        {
            throw new UndefinedStyleException("The default style can not be removed");
        }
        else
        {
            if (onlyIfUnused == true)
            {
                boolean styleInUse = false;
                Iterator itr;
                Map.Entry<String, Cell> cell;
                for(int i = 0; i < this.worksheets.size(); i++)
                {
                    itr = this.worksheets.get(i).getCells().entrySet().iterator();
                    while (itr.hasNext())
                    {
                        cell = (Map.Entry<String, Cell>)itr.next();
                        if (cell.getValue().getCellStyle().getName().equals(styleName))
                        {
                            styleInUse = true;
                            break;
                        }
                    }
                    if (styleInUse == true)
                    {
                        break;
                    }
                }
                if (styleInUse == false)
                {
                    this.styles.remove(index);
                }
            }
            else
            {
                this.styles.remove(index);
            }
        }        
    }
    
    /**
     * Method to prepare the styles before saving the workbook. Don't use the method otherwise. Styles will be reordered and probably removed from the style sheet
     * @return Returns a sorted collection of styles
     * @throws UndefinedStyleException Thrown if an unreferenced style was in the style sheet
     */
    public StyleCollection reorganizeStyles()
    {
        Iterator itr;
        Map.Entry<String, Cell> cell;
        List<Border> tempBorders = new ArrayList<>();
        List<Fill> tempFills = new ArrayList<>();
        List<Font> tempFonts = new ArrayList<>();
        List<NumberFormat> tempNumberFormats = new ArrayList<>();
        List<CellXf> tempCellXfs = new ArrayList<>();
        Style dateStyle = addStyle(BasicStyles.DateFormat(), true);
        int existingIndex = 0;
        boolean existing;
        int customNumberFormat = NumberFormat.CUSTOMFORMAT_START_NUMBER;
        for(int i = 0; i < this.styles.size(); i++)
        {
            this.styles.get(i).setInternalID(i);
            existing = false;
            for(int j = 0; j < tempBorders.size();j++)// item in tempBorders)
            {
                if (tempBorders.get(j).equals(this.styles.get(i).getCurrentBorder()) == true)
                {
                    existing = true;
                    existingIndex = tempBorders.get(j).getInternalID();
                    break;
                }
            }
            if (existing == false)
            {
                this.styles.get(i).getCurrentBorder().setInternalID(tempBorders.size());
                tempBorders.add(this.styles.get(i).getCurrentBorder());
            }
            else
            {
                this.styles.get(i).getCurrentBorder().setInternalID(existingIndex);
            }
            existing = false;
            for(int j = 0; j < tempFills.size();j++)
            {
                if (tempFills.get(j).equals(this.styles.get(i).getCurrentFill()) == true)
                {
                    existing = true;
                    existingIndex = tempFills.get(j).getInternalID();
                    break;
                }
            }
            if (existing == false)
            {
                this.styles.get(i).getCurrentFill().setInternalID(tempFills.size());
                tempFills.add(this.styles.get(i).getCurrentFill());
            }
            else
            {
                this.styles.get(i).getCurrentFill().setInternalID(existingIndex);
            }
            existing = false;
            for(int j = 0; j < tempFonts.size();j++)
            {
                if (tempFonts.get(j).equals(this.styles.get(i).getCurrentFont()) == true)
                {
                    existing = true;
                    existingIndex = tempFonts.get(j).getInternalID();
                    break;
                }
            }
            if (existing == false)
            {
                this.styles.get(i).getCurrentFont().setInternalID(tempFonts.size());
                tempFonts.add(this.styles.get(i).getCurrentFont());
            }
            else
            {
                this.styles.get(i).getCurrentFont().setInternalID(existingIndex);
            }
            existing = false;
            for(int j = 0; j < tempNumberFormats.size();j++)
            {
                if (tempNumberFormats.get(j).equals(this.styles.get(i).getCurrentNumberFormat()) == true)
                {
                    existing = true;
                    existingIndex = tempNumberFormats.get(j).getInternalID();
                    break;
                }
            }
            if (existing == false)
            {
                this.styles.get(i).getCurrentNumberFormat().setInternalID(tempNumberFormats.size());
                tempNumberFormats.add(this.styles.get(i).getCurrentNumberFormat());
            }
            else
            {
                this.styles.get(i).getCurrentNumberFormat().setInternalID(existingIndex);
            }
            if (this.styles.get(i).getCurrentNumberFormat().isCustomFormat() == true)
            {
                this.styles.get(i).getCurrentNumberFormat().setCustomFormatID(customNumberFormat);
                customNumberFormat++;
            }
            existing = false;
            for(int j = 0; j < tempCellXfs.size();j++)
            {
                if (tempCellXfs.get(j).equals(this.styles.get(i).getCurrentCellXf()) == true)
                {
                    existing = true;
                    existingIndex = tempCellXfs.get(j).getInternalID();
                    break;
                }
            }
            if (existing == false)
            {
                this.styles.get(i).getCurrentCellXf().setInternalID(tempCellXfs.size());
                tempCellXfs.add(this.styles.get(i).getCurrentCellXf());
            }
            else
            {
                this.styles.get(i).getCurrentCellXf().setInternalID(existingIndex);
            }
        }
        Style combiation;
        for(int j = 0; j < this.worksheets.size();j++)
        {
            //for(int k = 0; k < this.worksheets.get(j).getCells().size();k++)//KeyValuePair<string, Cell> cell in sheet.Cells)
            itr = this.worksheets.get(j).getCells().entrySet().iterator();
            while (itr.hasNext())
            {
                cell = (Map.Entry<String, Cell>)itr.next();
                if (cell.getValue().getFieldType() == Cell.CellType.DATE)
                {
                    if (cell.getValue().getCellStyle() == null)
                    {
                        combiation = dateStyle;
                    }
                    else
                    {
                        combiation = cell.getValue().getCellStyle().copy(dateStyle.getCurrentNumberFormat());
                    }
                    this.worksheets.get(j).getCells().get(cell.getKey()).setStyle(combiation, this);
                }
            }
        }

        Collections.sort(this.styles);
        Collections.sort(tempBorders);
        Collections.sort(tempFills);
        Collections.sort(tempFonts);
        Collections.sort(tempNumberFormats);
        Collections.sort(tempCellXfs);
        StyleCollection output = new StyleCollection();
        output.setBorders(tempBorders);
        output.setFonts(tempFonts);
        output.setFills(tempFills);
        output.setCellXfs(tempCellXfs);
        output.setNumberFormats(tempNumberFormats);
        return output;

    }
    
    /**
     * Method to resolve all merged cells in all worksheets. Only the value of the very first cell of the locked cells range will be visible. The other values are still present (set to EMPTY) but will not be stored in the worksheet.
     * @throws UndefinedStyleException Thrown if an unreferenced style was in the style sheet
     * @throws UnknownRangeException Thrown if the cell range was not found
     */
    public void resolveMergedCells()
    {
        Style mergStyle = BasicStyles.MergeCellStyle();
        int pos;
        List<Cell.Address> addresses;
        Cell cell;
        Worksheet sheet;
        Cell.Address address;
        Iterator itr;
        Map.Entry<String, Cell.Range> range;
        for (int i = 0; i < this.worksheets.size(); i++)
        {
            sheet = this.worksheets.get(i);
            itr = sheet.getMergedCells().entrySet().iterator();
            while (itr.hasNext())
            {
                range = (Map.Entry<String, Cell.Range>)itr.next();
                pos = 0;
                addresses = Cell.getCellRange(range.getValue().StartAddress, range.getValue().EndAddress);
                for (int j = 0; j < addresses.size(); j++)
                {
                    address = addresses.get(j);
                    if (sheet.getCells().containsKey(address.toString()) == false)
                    {
                        cell = new Cell();
                        cell.setFieldType(Cell.CellType.EMPTY);
                        cell.setRowAddress(address.Row);
                        cell.setColumnAddress(address.Column);
                        sheet.addCell(cell);
                    }
                    else
                    {
                        cell = sheet.getCells().get(address.toString());
                    }
                    if (pos != 0)
                    {
                        cell.setFieldType(Cell.CellType.EMPTY);
                    }
                    cell.setStyle(mergStyle, this);
                    pos++;
                }
            }
        }
    }    

    /**
    * Saves the workbook
    * @throws IOException Throws IOException in case of an error
    */
    public void save() throws IOException
    {
        LowLevel l = new LowLevel(this);
        l.save();
    }
    
    /**
    * Saves the workbook with the defined name
    * @param filename filename of the saved workbook
    * @throws IOException Thrown in case of an error
    */
    public void saveAs(String filename) throws IOException
    {
        String backup = this.filename;
        this.filename = filename;
        LowLevel l = new LowLevel(this);
        l.save();
        this.filename = backup;
    }    
    
        
}
