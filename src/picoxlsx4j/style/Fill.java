/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2016
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.style;

/**
 *
 * @author Raphael Stoeckli
 */
public class Fill implements Comparable<Fill> {
    
    public static final String DEFAULTCOLOR = "FF000000";
    
    public enum PatternValue
    {
        /**
         * No pattern (default)
         */
        none,
        /**
        * Solid fill (for colors)
        */
        solid,
        /**
        * Dark gray fill
        */
        darkGray,
        /**
        * Medium gray fill
        */
        mediumGray,
        /**
        * Light gray fill
        */
        lightGray,
        /**
        * 6.25% gray fill
        */
        gray0625,
        /**
        * 12.5% gray fill
        */
        gray125,
    }
    
    public enum FillType
    {
        /**
        * Color defines a pattern color
        */
        patternColor,
        /**
         * Color defines a solid fill color
        */
        fillColor,
    }    
    
    /**
     * Gets the pattern name from the enum
     * @param pattern Enum to process
     * @return The valid value of the pattern as String
     */
    public static String getPatternName(PatternValue pattern)
    {
        String output = "";
        switch (pattern)
        {
            case none:
                output = "none";
                break;
            case solid:
                output = "solid";
                break;
            case darkGray:
                output = "darkGray";
                break;
            case mediumGray:
                output = "mediumGray";
                break;
            case lightGray:
                output = "lightGray";
                break;
            case gray0625:
                output = "gray0625";
                break;
            case gray125:
                output = "gray125";
                break;
            default:
                output = "none";
                break;
        }
        return output;
    }    

    public int indexedColor;
    public PatternValue patternFill;
    public String foregroundColor;
    public String backgroundColor;
    public int internalID;    

    /**
     * Gets the indexed color (Default is 64)
     * @return Indexed color
     */
    public int getIndexedColor() {
        return indexedColor;
    }

    /**
     * Sets the indexed color (Default is 64)
     * @param indexedColor Indexed color
     */
    public void setIndexedColor(int indexedColor) {
        this.indexedColor = indexedColor;
    }

    /**
     * Gets the pattern type of the fill (Default is none)
     * @return Pattern type of the fill
     */
    public PatternValue getPatternFill() {
        return patternFill;
    }

    /**
     * Sets the pattern type of the fill (Default is none)
     * @param patternFill Pattern type of the fill
     */
    public void setPatternFill(PatternValue patternFill) {
        this.patternFill = patternFill;
    }

    /**
     * Gets the foreground color of the fill
     * @return Foreground color of the fill
     */
    public String getForegroundColor() {
        return foregroundColor;
    }

    /**
     * Sets the foreground color of the fill
     * @param foregroundColor Foreground color of the fill
     */
    public void setForegroundColor(String foregroundColor) {
        this.foregroundColor = foregroundColor;
    }

    /**
     * Gets the Background color of the fill
     * @return Background color of the fill
     */
    public String getBackgroundColor() {
        return backgroundColor;
    }

    /**
     * Sets the background color of the fill
     * @param backgroundColor Background color of the fill
     */
    public void setBackgroundColor(String backgroundColor) {
        this.backgroundColor = backgroundColor;
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
     * Default constructor
     */
    public Fill()
    {
        this.indexedColor = 64;
        this.patternFill = PatternValue.none;
        this.foregroundColor = DEFAULTCOLOR;
        this.backgroundColor = DEFAULTCOLOR;
    }  
    
    /**
     * Constructor with foreground and background color
     * @param forground Foreground color of the fill
     * @param background Background color of the fill
     */
    public Fill(String forground, String background)
    {
        this.backgroundColor = background;
        this.foregroundColor = forground;
        this.indexedColor = 64;
        this.patternFill = PatternValue.solid;
    }
    
    /**
     * Constructor with color value and fill type
     * @param value Color value
     * @param filltype Fill type (fill or pattern)
     */
    public Fill(String value, FillType filltype)
    {
        if (filltype == FillType.fillColor)
        {
            this.backgroundColor = value;
            this.foregroundColor = DEFAULTCOLOR;
        }
        else
        {
            this.backgroundColor = DEFAULTCOLOR;
            this.foregroundColor = value;
        }
        this.indexedColor = 64;
        this.patternFill = PatternValue.solid;
    }
    
    /**
     * Seth the color an the depending fill type
     * @param value Color value
     * @param filltype Fill type (fill or pattern)
     */
    public void SetColor(String value, FillType filltype)
    {
        if (filltype == FillType.fillColor)
        {
            this.foregroundColor = value;
            this.backgroundColor = DEFAULTCOLOR;
        }
        else
        {
            this.foregroundColor = DEFAULTCOLOR;
            this.backgroundColor = value;
        }
        this.patternFill = PatternValue.solid;
    }
    
    /**
     * Method to compare two objects for sorting purpose
     * @param o Other object to compare with this object
     * @return True if both objects are equal, otherwise false
     */    
    @Override
    public boolean equals(Object o)
    {
        if (o == null) { return false; }
        Fill other = (Fill)o;
        if (this.indexedColor != other.getIndexedColor()) { return false; }
        if (this.patternFill != other.getPatternFill()) { return false; }
        if (this.foregroundColor.equals(other.getForegroundColor()) == false) { return false; }
        if (this.backgroundColor.equals(other.getBackgroundColor()) == false) { return false; }
        else { return true; }
    }    
     
    /**
     * Method to compare two objects for sorting purpose
     * @param o Other object to compare with this object
     * @return -1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.
     */        
    @Override
    public int compareTo(Fill o) {
        return Integer.compare(internalID, o.getInternalID());
    }
    
    /**
     * Method to copy the current object to a new one
     * @return Copy of the current object without the internal ID
     */       
    public Fill copy()
    {
        Fill copy = new Fill();
        copy.setBackgroundColor(this.backgroundColor);
        copy.setForegroundColor(this.foregroundColor);
        copy.setIndexedColor(this.indexedColor);
        copy.setPatternFill(this.patternFill);
        return copy;
    }    
    
}
