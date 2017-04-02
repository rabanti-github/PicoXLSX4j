/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2017
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.style;

import java.security.SecureRandom;

/**
 * Class representing a Style with references to sub-styles within a style sheet. An instance of this class is only a container for the different sub-styles. These sub-styles contains the actual styling information.
 * @author Raphael Stoeckli
 */
public class Style  implements Comparable<Style>{
    
    private Font currentFont;
    private Fill currentFill;
    private Border currentBorder;
    private NumberFormat currentNumberFormat;
    private CellXf currentCellXf;
    private int internalID;
    private String name;

    /**
     * Gets the Current font object of the style
     * @return Current Font object of the style
     */
    public Font getCurrentFont() {
        return currentFont;
    }

    /**
     * Sets the current Font object of the style
     * @param currentFont Current Font object of the style
     */
    public void setCurrentFont(Font currentFont) {
        this.currentFont = currentFont;
    }

    /**
     * Gets the current Fill object of the style
     * @return Current Fill object of the style
     */
    public Fill getCurrentFill() {
        return currentFill;
    }

    /**
     * Sets the current Fill object of the style
     * @param currentFill Current Fill object of the style
     */
    public void setCurrentFill(Fill currentFill) {
        this.currentFill = currentFill;
    }

    /**
     * Gets the current Border object of the style
     * @return Current Border object of the style
     */
    public Border getCurrentBorder() {
        return currentBorder;
    }

    /**
     * Sets the current Border object of the style
     * @param currentBorder Current Border object of the style
     */
    public void setCurrentBorder(Border currentBorder) {
        this.currentBorder = currentBorder;
    }

    /**
     * Gets the current NumberFormat object of the style
     * @return Current NumberFormat object of the style
     */
    public NumberFormat getCurrentNumberFormat() {
        return currentNumberFormat;
    }

    /**
     * Sets the current NumberFormat object of the style
     * @param currentNumberFormat Current NumberFormat object of the style
     */
    public void setCurrentNumberFormat(NumberFormat currentNumberFormat) {
        this.currentNumberFormat = currentNumberFormat;
    }

    /**
     * Gets the current CellXf object of the style
     * @return Current CellXf object of the style
     */
    public CellXf getCurrentCellXf() {
        return currentCellXf;
    }

    /**
     * Sets the current CellXf object of the style
     * @param currentCellXf Current CellXf object of the style
     */
    public void setCurrentCellXf(CellXf currentCellXf) {
        this.currentCellXf = currentCellXf;
    }

    /**
     * Gets the internal ID for sorting purpose
     * @return Internal ID
     */
    public int getInternalID() {
        return internalID;
    }

    /**
     * Sets the internal ID for sorting purpose
     * @param internalID Internal ID
     */
    public void setInternalID(int internalID) {
        this.internalID = internalID;
    }
    

    /**
     * Gets the name of the style. If not defined, a random name will be generated when the style is created
     * @return Name of the style
     */
    public String getName() {
        return name;
    }

    /**
     * Sets the Name of the style
     * @param name Name of the style
     */
    public void setName(String name) {
        this.name = name;
    }
   
    /**
     * Default constructor
     */
   public Style ()
   {
      init(); 
   }    
    
   /**
    * Constructor with definition of the style name
    * @param name Name of the style
    */
   public Style (String name)
   {
       init();
       this.name = name;
   }
   
   /**
    * Init method for the constructors
    */
   private void init()
   {
        this.currentFont = new Font();
        this.currentFill = new Fill();
        this.currentBorder = new Border();
        this.currentNumberFormat = new NumberFormat();
        this.currentCellXf = new CellXf();
        this.name = createUniqueName();
   }
   
   /**
    * Method to determine the equality of two objects
    * @param o Object to compare against this object
    * @return True if both objects are equal, otherwise false
    */
    @Override
    public boolean equals(Object o)
    {
        if (o == null) {return false;}
        Style other = (Style)o;
        if (this.currentBorder != null && other.getCurrentBorder() != null)
        {
            if (this.currentBorder.equals(other.getCurrentBorder()) == false) { return false; }
        }
        if ((this.currentBorder == null || other.getCurrentBorder() == null) && !(this.currentBorder == null && other.getCurrentBorder() == null)) { return false; }

        if (this.currentFill != null && other.getCurrentFill() != null)
        {
            if (this.currentFill.equals(other.getCurrentFill()) == false) { return false; }
        }
        if ((this.currentFill == null || other.getCurrentFill() == null) && !(this.currentFill == null && other.getCurrentFill() == null)) { return false; }

        if (this.currentFont != null && other.getCurrentFont() != null)
        {
            if (this.currentFont.equals(other.getCurrentFont()) == false) { return false; }
        }
        if ((this.currentFont == null || other.getCurrentFont() == null) && !(this.currentFont == null && other.getCurrentFont() == null)) { return false; }

        if (this.currentNumberFormat != null && other.getCurrentNumberFormat() != null)
        {
            if (this.currentNumberFormat.equals(other.getCurrentNumberFormat()) == false) { return false; }
        }
        if ((this.currentNumberFormat == null || other.getCurrentNumberFormat() == null) && !(this.currentNumberFormat == null && other.getCurrentNumberFormat() == null)) { return false; }

        if (this.currentCellXf != null && other.getCurrentCellXf() != null)
        {
            if (this.currentCellXf.equals(other.getCurrentCellXf()) == false) { return false; }
        }
        if ((this.currentCellXf == null || other.getCurrentCellXf() == null) && !(this.currentCellXf == null && other.getCurrentCellXf() == null)) { return false; }
        return true;
    }

    /**
     * Method to compare two objects for sorting purpose
     * @param o Other object to compare with this object
     * @return -1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.
     */         
    @Override
    public int compareTo(Style o) {
        return Integer.compare(internalID, o.getInternalID());
    }
 
    /**
     * Method to copy the current object to a new one
     * @return Copy of the current object without the internal ID
     */    
    public Style copy()
    {
        Style copy = null;
        try
        {
        copy = new Style(this.name + "_copy");
        copy.setCurrentBorder(this.currentBorder.copy());
        copy.setCurrentFill(this.currentFill.copy());
        copy.setCurrentFont(this.currentFont.copy());
        copy.setCurrentNumberFormat(this.currentNumberFormat.copy());
        copy.setCurrentCellXf(this.currentCellXf.copy());
        }
        catch (Exception e)
        {
            // Will never happen, because earlier checked
        }
        return copy;
    }    
   
    /**
     * Method to copy the current object to a new one
     * @param overwriteFormat NumberFormat object to replace the original object of the current object in the copy
     * @return Copy of the current object without the internal ID
     */
    public Style copy(NumberFormat overwriteFormat) 
    {
        Style copy = null;
        try
        {
        copy = new Style(this.name + "_copy");
        copy.setCurrentBorder(this.currentBorder.copy());
        copy.setCurrentFill(this.currentFill.copy());
        copy.setCurrentFont(this.currentFont.copy());
        copy.setCurrentNumberFormat(overwriteFormat);
        copy.setCurrentCellXf(this.currentCellXf.copy());
        }
        catch (Exception e)
        {
            // Will never happen, because earlier checked
        }
        return copy;
    }  
    
    /**
     * Creates a random style names using a Crypto Service Provider (prevents same random numbers due to too fast processing)
     * @return Random style name
     */
    private String createUniqueName()
    {
      SecureRandom random = new SecureRandom();
      byte bytes[] = new byte[20];
      int number = random.nextInt(Integer.MAX_VALUE);
      int number2 = random.nextInt(Integer.MAX_VALUE);
      return "Style" + Integer.toString(number) + "-" + Integer.toString(number2);
    }
    
}
