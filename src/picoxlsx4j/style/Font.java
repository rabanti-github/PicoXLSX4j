/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.style;

/**
 * Class representing a Font entry. The Font entry is used to define text formatting
 * @author Raphael Stoeckli
 */
public class Font implements Comparable<Font> {

    /**
     * Default font family as constant
     */
    public static final String DEFAULTFONT = "Calibri";
    
    /**
     * Enum for the vertical alignment of the text from base line
     */
    public enum VerticalAlignValue
    {
        // baseline, // Maybe not used in Excel
        /**
         * Text will be rendered as subscript
         */
        subscript,
        /**
         * Text will be rendered as superscript
         */
        superscript,
        /**
         * Text will be rendered normal
         */
        none,
    }    
    
    public enum SchemeValue
    {
        /**
         * Font scheme is major
         */
        major,
        /**
         * Font scheme is minor (default)
         */
        minor,
        /**
         * No Font scheme is used
         */
        none,
    }    
    
    private int size;
    private String name;
    private String family;
    private int colorTheme;
    private String colorValue;
    private SchemeValue scheme;
    private VerticalAlignValue verticalAlign;
    private boolean bold;
    private boolean italic;
    private boolean underline;
    private boolean doubleUnderline;
    private boolean strike;
    private String charset;
    private int internalID;

    /**
     * Gets the font size. Valid range is from 8 to 75
     * @return Font size
     */
    public int getSize() {
        return size;
    }

    /**
     * Sets the Font size. Valid range is from 8 to 75
     * @param size Font size
     */
    public void setSize(int size) {
        if (size < 8) { this.size = 8; }
        else if (size > 75) { this.size = 72; }
        else { this.size = size; }
    }

    /**
     * Gets the font name (Default is Calibri)
     * @return Font name
     */
    public String getName() {
        return name;
    }

    /**
     * Sets the font name (Default is Calibri)
     * @param name Font name
     */
    public void setName(String name) {
        this.name = name;
    }

    /**
     * Gets the font family (Default is 2)
     * @return Font family
     */
    public String getFamily() {
        return family;
    }

    /**
     * Sets the font family (Default is 2)
     * @param family Font family
     */
    public void setFamily(String family) {
        this.family = family;
    }

    /**
     * Gets the font color theme (Default is 1)
     * @return Font color theme
     */
    public int getColorTheme() {
        return colorTheme;
    }

    /**
     * Sets the font color theme (Default is 1)
     * @param colorTheme Font color theme
     */
    public void setColorTheme(int colorTheme) {
        this.colorTheme = colorTheme;
    }

    /**
     * Gets the Font color (default is empty)
     * @return Font color
     */
    public String getColorValue() {
        return colorValue;
    }

    /**
     * Sets the font color (default is empty)
     * @param colorValue Font color
     */
    public void setColorValue(String colorValue) {
        this.colorValue = colorValue;
    }

    /**
     * Gets the font scheme (Default is minor)
     * @return Font scheme
     */
    public SchemeValue getScheme() {
        return scheme;
    }

    /**
     * Sets the Font scheme (Default is minor)
     * @param scheme Font scheme
     */
    public void setScheme(SchemeValue scheme) {
        this.scheme = scheme;
    }

    /**
     * Gets the alignment of the font (Default is none)
     * @return Alignment of the font
     */
    public VerticalAlignValue getVerticalAlign() {
        return verticalAlign;
    }

    /**
     * Sets the Alignment of the font (Default is none)
     * @param verticalAlign Alignment of the font
     */
    public void setVerticalAlign(VerticalAlignValue verticalAlign) {
        this.verticalAlign = verticalAlign;
    }

    /**
     * Gets the bold parameter of the font
     * @return If true, the font is bold
     */
    public boolean isBold() {
        return bold;
    }

    /**
     * Sets the bold parameter of the font
     * @param bold If true, the font is bold
     */
    public void setBold(boolean bold) {
        this.bold = bold;
    }

    /**
     * Gets the italic parameter of the font
     * @return If true, the font is italic
     */
    public boolean isItalic() {
        return italic;
    }

    /**
     * Sets the italic parameter of the font
     * @param italic If true, the font is italic
     */
    public void setItalic(boolean italic) {
        this.italic = italic;
    }

    /**
     * Gets the underline parameter of the font
     * @return If true, the font as one underline
     */
    public boolean isUnderline() {
        return underline;
    }

    /**
     * Sets the underline parameter of the font
     * @param underline If true, the font as one underline
     */
    public void setUnderline(boolean underline) {
        this.underline = underline;
    }

    /**
     * Gets the double-underline parameter of the font
     * @return If true, the font ha a double underline
     */
    public boolean isDoubleUnderline() {
        return doubleUnderline;
    }

    /**
     * Sets the double-underline parameter of the font
     * @param doubleUnderline If true, the font ha a double underline
     */
    public void setDoubleUnderline(boolean doubleUnderline) {
        this.doubleUnderline = doubleUnderline;
    }

    /**
     * Gets the strike parameter of the font
     * @return If true, the font is stroked through
     */
    public boolean isStrike() {
        return strike;
    }

    /**
     * Sets the strike parameter of the font
     * @param strike If true, the font is stroked through
     */
    public void setStrike(boolean strike) {
        this.strike = strike;
    }

    /**
     * Gets the charset of the Font (Default is empty)
     * @return Charset of the Font
     */
    public String getCharset() {
        return charset;
    }

    /**
     * Sets the charset of the Font (Default is empty)
     * @param charset Charset of the Font
     */
    public void setCharset(String charset) {
        this.charset = charset;
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
     * Gets whether this object is the default font
     * @return In true the font is equals the default font
     */
    public boolean isDefaultFont()
    {
        Font temp = new Font();
        return this.equals(temp); 
    }
    
    /**
     * Default constructor
     */
    public Font()
    {
        this.size = 11;
        this.name = DEFAULTFONT;
        this.family = "2";
        this.colorTheme = 1;
        this.colorValue = "";
        this.charset = "";
        this.scheme = SchemeValue.minor;
        this.verticalAlign = VerticalAlignValue.none;
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
        Font other = (Font)o;
        if (this.bold != other.isBold()) { return false; }
        if (this.colorTheme != other.getColorTheme()) { return false; }
        if (this.doubleUnderline != other.isDoubleUnderline()) { return false; }
        if (this.family.equals(other.getFamily()) == false) { return false; }
        if (this.italic != other.isItalic()) { return false; }
        if (this.name.equals(other.getName()) == false) { return false; }
        if (this.scheme != other.getScheme()) { return false; }
        if (this.verticalAlign != other.getVerticalAlign()) { return false; }
        if (this.charset.equals(other.getCharset()) == false) { return false; }
        if (this.size != other.getSize()) { return false; }
        if (this.strike != other.isStrike()) { return false; }
        if (this.underline != other.isUnderline()) { return false; }
        else { return true; }
    }    
    
    /**
     * Method to compare two objects for sorting purpose
     * @param o Other object to compare with this object
     * @return -1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.
     */      
    @Override
    public int compareTo(Font o) {
        return Integer.compare(internalID, o.getInternalID());
    }
    
    /**
     * Method to copy the current object to a new one
     * @return Copy of the current object without the internal ID
     */          
    public Font copy()
    {
        Font copy = new Font();
        copy.setBold(this.bold);
        copy.setCharset(this.charset);
        copy.setColorTheme(this.colorTheme);
        copy.setVerticalAlign(this.verticalAlign);
        copy.setDoubleUnderline(this.doubleUnderline);
        copy.setFamily(this.family);
        copy.setItalic(this.italic);
        copy.setName(this.name);
        copy.setScheme(this.scheme);
        copy.setSize(this.size);
        copy.setStrike(this.strike);
        copy.setUnderline(this.underline);
        return copy;
    }    
    
}
