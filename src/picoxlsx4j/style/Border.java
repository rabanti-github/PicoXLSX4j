/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2017
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.style;

/**
 * Class representing a Border entry. The Border entry is used to define frames and cell borders
 * @author Raphael Stoeckli
 */
public class Border implements Comparable<Border> {

    /**
     * Enum for the border style
     */
    public enum StyleValue
    {
        /**
        * no border
        */
        none,
        /**
        * hair border
        */
        hair,
        /**
        * dotted border
        */
        dotted,
        /**
        * dashed border with double-dots
        */
        dashDotDot,
        /**
        * dash-dotted border
        */
        dashDot,
        /**
        * dashed border
        */
        dashed,
        /**
        * thin border
        */
        thin,
        /**
        * medium-dashed border with double-dots
        */
        mediumDashDotDot,
        /**
        * slant dash-dotted border
        */
        slantDashDot,
        /**
        * medium dash-dotted border
        */
        mediumDashDot,
        /**
        * medium dashed border
        */
        mediumDashed,
        /**
        * medium border
        */
        medium,
        /**
        * thick border
        */
        thick,
        /**
        * double border
        */
        s_double,
    }
    
    /**
     * Gets the border style name from the enum
     * @param style Enum to process
     * @return The valid value of the border style as String
     */
    public static String getStyleName(StyleValue style)
    {
        String output = "";
        switch (style)
        {
            case none:
                output = "";
                break;
            case hair:
                break;
            case dotted:
                output = "dotted";
                break;
            case dashDotDot:
                output = "dashDotDot";
                break;
            case dashDot:
                output = "dashDot";
                break;
            case dashed:
                output = "dashed";
                break;
            case thin:
                output = "thin";
                break;
            case mediumDashDotDot:
                output = "mediumDashDotDot";
                break;
            case slantDashDot:
                output = "slantDashDot";
                break;
            case mediumDashDot:
                output = "mediumDashDot";
                break;
            case mediumDashed:
                output = "mediumDashed";
                break;
            case medium:
                output = "medium";
                break;
            case thick:
                output = "thick";
                break;
            case s_double:
                output = "double";
                break;
            default:
                output = "";
                break;
        }
        return output;
    }    
    
    private StyleValue leftStyle;
    private StyleValue rightStyle;
    private StyleValue topStyle;
    private StyleValue bottomStyle;
    private StyleValue diagonalStyle;
    private boolean diagonalDown;
    private boolean diagonalUp;
    private String leftColor;
    private String rightColor;
    private String topColor;
    private String bottomColor;
    private String diagonalColor;
    private int internalID;    

    /**
     * Gets the style of left cell border
     * @return Style of left cell border
     */
    public StyleValue getLeftStyle() {
        return leftStyle;
    }

    /**
     * Sets the style of left cell border
     * @param leftStyle Style of left cell border
     */
    public void setLeftStyle(StyleValue leftStyle) {
        this.leftStyle = leftStyle;
    }

    /**
     * Gets the style of right cell border
     * @return Style of right cell border
     */
    public StyleValue getRightStyle() {
        return rightStyle;
    }

    /**
     * Sets the style of right cell border
     * @param rightStyle Style of right cell border
     */
    public void setRightStyle(StyleValue rightStyle) {
        this.rightStyle = rightStyle;
    }

    /**
     * Gets the style of top cell border
     * @return Style of top cell border
     */
    public StyleValue getTopStyle() {
        return topStyle;
    }

    /**
     * Sets the style of top cell border
     * @param topStyle Style of top cell border
     */
    public void setTopStyle(StyleValue topStyle) {
        this.topStyle = topStyle;
    }

    /**
     * Gets the style of bottom cell border
     * @return Style of bottom cell border
     */
    public StyleValue getBottomStyle() {
        return bottomStyle;
    }

    /**
     * Sets the style of bottom cell border
     * @param bottomStyle Style of bottom cell border
     */
    public void setBottomStyle(StyleValue bottomStyle) {
        this.bottomStyle = bottomStyle;
    }

    /**
     * Gets the style of the diagonal lines
     * @return Style of the diagonal lines
     */
    public StyleValue getDiagonalStyle() {
        return diagonalStyle;
    }

    /**
     * Sets the style of the diagonal lines
     * @param diagonalStyle Style of the diagonal lines
     */
    public void setDiagonalStyle(StyleValue diagonalStyle) {
        this.diagonalStyle = diagonalStyle;
    }

    /**
     * Gets the downwards diagonal line
     * @return If true, the downwards diagonal line is used
     */
    public boolean isDiagonalDown() {
        return diagonalDown;
    }

    /**
     * Sets the downwards diagonal line
     * @param diagonalDown If true, the downwards diagonal line is used
     */
    public void setDiagonalDown(boolean diagonalDown) {
        this.diagonalDown = diagonalDown;
    }

    /**
     * Gets the upwards diagonal line
     * @return If true, the upwards diagonal line is used
     */
    public boolean isDiagonalUp() {
        return diagonalUp;
    }

    /**
     * Sets the upwards diagonal line
     * @param diagonalUp If true, the upwards diagonal line is used
     */
    public void setDiagonalUp(boolean diagonalUp) {
        this.diagonalUp = diagonalUp;
    }

    /**
     * Gets the color code (ARGB) of the left border
     * @return Color code (ARGB)
     */
    public String getLeftColor() {
        return leftColor;
    }

    /**
     * Sets the color code (ARGB) of the left border
     * @param leftColor Color code (ARGB)
     */
    public void setLeftColor(String leftColor) {
        this.leftColor = leftColor;
    }

    /**
     * Gets the color code (ARGB) of the right border
     * @return Color code (ARGB)
     */
    public String getRightColor() {
        return rightColor;
    }

    /**
     * Sets the color code (ARGB) of the right border
     * @param rightColor Color code (ARGB)
     */
    public void setRightColor(String rightColor) {
        this.rightColor = rightColor;
    }

    /**
     * Gets the color code (ARGB) of the top border
     * @return Color code (ARGB)
     */
    public String getTopColor() {
        return topColor;
    }

    /**
     * Sets the color code (ARGB) of the top border
     * @param topColor Color code (ARGB)
     */
    public void setTopColor(String topColor) {
        this.topColor = topColor;
    }

    /**
     * Gets the color code (ARGB) of the bottom border
     * @return Color code (ARGB)
     */
    public String getBottomColor() {
        return bottomColor;
    }

    /**
     * Sets the color code (ARGB) of the bottom border
     * @param bottomColor Color code (ARGB)
     */
    public void setBottomColor(String bottomColor) {
        this.bottomColor = bottomColor;
    }

    /**
     * Gets the color code (ARGB) of the diagonal lines
     * @return Color code (ARGB)
     */
    public String getDiagonalColor() {
        return diagonalColor;
    }

    /**
     * Sets the color code (ARGB) of the diagonal lines
     * @param diagonalColor Color code (ARGB)
     */
    public void setDiagonalColor(String diagonalColor) {
        this.diagonalColor = diagonalColor;
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
    public Border()
    {
        this.bottomColor = "";
        this.topColor = "";
        this.leftColor = "";
        this.rightColor = "";
        this.diagonalColor = "";
        this.leftStyle = StyleValue.none;
        this.rightStyle = StyleValue.none;
        this.topStyle = StyleValue.none;
        this.bottomStyle = StyleValue.none;
        this.diagonalStyle = StyleValue.none;
        this.diagonalDown = false;
        this.diagonalUp = false;
    }    
    
    /**
     * Method to compare two objects for sorting purpose
     * @param o Other object to compare with this object
     * @return True if both objects are equal, otherwise false
     */
    public boolean equals(Object o)
    {
        if (o == null) { return false; }
        Border other = (Border)o;
        if (this.bottomColor.equals(other.getBottomColor()) == false) { return false; }
        if (this.bottomStyle != other.getBottomStyle()) { return false; }
        if (this.diagonalColor.equals(other.getDiagonalColor()) == false) { return false; }
        if (this.diagonalDown != other.isDiagonalDown()) { return false; }
        if (this.diagonalStyle != other.getDiagonalStyle()) { return false; }
        if (this.diagonalUp != other.isDiagonalUp()) { return false; }
        if (this.leftColor.equals(other.getLeftColor()) == false) { return false; }
        if (this.leftStyle != other.getLeftStyle()) { return false; }
        if (this.rightColor.equals(other.getRightColor()) == false) { return false; }
        if (this.rightStyle != other.getRightStyle()) { return false; }
        if (this.topColor.equals(other.getTopColor()) == false) { return false; }
        if (this.topStyle != other.getTopStyle()) { return false; }
        else { return true; }
    }    
    
    /**
     * Method to determine whether the object has no values but the default values (means: is empty and must not be processed)
     * @return True if empty, otherwise false
     */
    public boolean isEmpty()
    {
        boolean state = true;
        if (this.bottomColor.length() == 0) {state = false;}
        if (this.topColor.length() == 0) {state = false;}
        if (this.leftColor.length() == 0) {state = false;}
        if (this.rightColor.length() == 0) {state = false;}
        if (this.diagonalColor.length() == 0) {state = false;}
        if (this.leftStyle != StyleValue.none) {state = false;}
        if (this.rightStyle != StyleValue.none) {state = false;}
        if (this.topStyle != StyleValue.none) {state = false;}
        if (this.bottomStyle != StyleValue.none) {state = false;}
        if (this.diagonalStyle != StyleValue.none) {state = false;}
        if (this.diagonalDown != false) {state = false;}
        if (this.diagonalUp != false) { state = false; }
        return state;
    }    
    
    /**
     * Method to compare two objects for sorting purpose
     * @param o Other object to compare with this object
     * @return -1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.
     */
    @Override
    public int compareTo(Border o) {
        return Integer.compare(internalID, o.getInternalID());
    }
    
    /**
     * Method to copy the current object to a new one
     * @return Copy of the current object without the internal ID
     */
    public Border copy()
    {
        Border copy = new Border();
        copy.setBottomColor(this.bottomColor);
        copy.setBottomStyle(this.bottomStyle);
        copy.setDiagonalColor(this.diagonalColor);
        copy.setDiagonalDown(this.diagonalDown);
        copy.setDiagonalStyle(this.diagonalStyle);
        copy.setDiagonalUp(this.diagonalUp);
        copy.setLeftColor(this.leftColor);
        copy.setLeftStyle(this.leftStyle);
        copy.setRightColor(this.rightColor);
        copy.setRightStyle(this.rightStyle);
        copy.setTopColor(this.topColor);
        copy.setTopStyle(this.topStyle);
        return copy;
    }    
    
}
