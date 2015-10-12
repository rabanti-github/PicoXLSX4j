/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.style;

import picoxlsx4j.exception.OutOfRangeException;

/**
 * Class representing an XF entry. The XF entry is used to make reference to other style instances like Border oder Fill and for the positioning of the cell content
 * @author Raphael Stoeckli
 */
public class CellXf implements Comparable<CellXf>
{

    /**
     * Enum for the horizontal alignment of a cell 
     */
     public enum HorizontalAlignValue
     {
         /**
         * Content will be aligned left
         */
         left,
         /**
         * Content will be aligned in the center
         */
         center,
         /**
         * Content will be aligned right
         */
         right,
         /**
         * Content will fill up the cell
         */
         fill,
         /**
         * justify alignment
         */
         justify,
         /**
         * General alignment
         */
         general,
         /**
         * Center continuous alignment
         */
         centerContinuous,
         /**
         * Distributed alignment
         */
         distributed,
         /**
         * No alignment. The alignment will not be used in a style
         */
         none,
     };

     /**
      * Enum for the vertical alignment of a cell 
      */
     public enum VerticallAlignValue
     {
         /**
         * Content will be aligned on the bottom (default)
         */
         bottom,
         /**
         * Content will be aligned on the top
         */
         top,
         /**
         * Content will be aligned in the center
         */
         center,
         /**
         * justify alignment
         */
         justify,
         /**
         * Distributed alignment
         */
         distributed,
         /**
         * No alignment. The alignment will not be used in a style
         */
         none,
     }     

     /**
      * Enum for text break options
      */
     public enum TextBreakValue
     {
         /**
          * Word wrap is active
          */
         wrapText,
         /**
          * Text will be resized to fit the cell
          */
         shrinkToFit,
         /**
          * Text will overflow in cell
          */
         none,
     }

     /**
      * Enum for the general text alignment direction
      */
     public enum TextDirectionValue
     {
         /**
          * Text direction is horizontal (default)
          */
         horizontal,
         /**
          * Text direction is vertical
          */
         vertical,
     }

     private int textRotation;
     private TextDirectionValue textDirection;           
     private HorizontalAlignValue horizontalAlign;
     private VerticallAlignValue verticalAlign;
     private TextBreakValue alignment;
     private int internalID;

     /**
      * Gets the text rotation in degrees (from +90 to -90)
      * @return Text rotation in degrees (from +90 to -90)
      */
     public int getTextRotation() {
         return textRotation;
     }

     /**
      * Sets the text rotation in degrees (from +90 to -90)
      * @param textRotation Text rotation in degrees (from +90 to -90)
      * @throws OutOfRangeException Thrown if the rotation angle is out of range
      */
     public void setTextRotation(int textRotation) throws OutOfRangeException {
         this.textRotation = textRotation;
         this.textDirection = TextDirectionValue.horizontal;
         calculateInternalRotation();
     }

     /**
      * Gets the direction of the text within the cell
      * @return Direction of the text within the cell
      */
     public TextDirectionValue getTextDirection() {
         return textDirection;
     }

     /**
      * Sets the direction of the text within the cell
      * @param textDirection Direction of the text within the cell
      * @throws OutOfRangeException Thrown if the text rotation and direction causes a conflict
      */
     public void setTextDirection(TextDirectionValue textDirection) throws OutOfRangeException {
         this.textDirection = textDirection;
         calculateInternalRotation();            
     }

     /**
      * Gets the horizontal alignment of the style
      * @return Horizontal alignment of the style
      */
     public HorizontalAlignValue getHorizontalAlign() {
         return horizontalAlign;
     }

     /**
      * Sets the horizontal alignment of the style
      * @param horizontalAlign Horizontal alignment of the style
      */
     public void setHorizontalAlign(HorizontalAlignValue horizontalAlign) {
         this.horizontalAlign = horizontalAlign;
     }

     /**
      * Gets the vertical alignment of the style
      * @return Vertical alignment of the style
      */
     public VerticallAlignValue getVerticalAlign() {
         return verticalAlign;
     }

     /**
      * Sets the vertical alignment of the style
      * @param verticalAlign Vertical alignment of the style
      */
     public void setVerticalAlign(VerticallAlignValue verticalAlign) {
         this.verticalAlign = verticalAlign;
     }

     /**
      * Gets the text break options of the style
      * @return Text break options of the style
      */
     public TextBreakValue getAlignment() {
         return alignment;
     }

     /**
      * Sets the text break options of the style
      * @param alignment Text break options of the style
      */
     public void setAlignment(TextBreakValue alignment) {
         this.alignment = alignment;
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
     public CellXf()
     {
         this.horizontalAlign = HorizontalAlignValue.none;
         this.alignment = TextBreakValue.none;
         this.textDirection = TextDirectionValue.horizontal;
         this.verticalAlign = VerticallAlignValue.none;
         this.textRotation = 0;            
     }

     /**
      * Method to calculate the internal text rotation. The text direction and rotation are handled internally by the text rotation value
      * @return Returns the valid rotation in degrees for internal uses (LowLevel)
      * @throws OutOfRangeException Thrown if the rotation is out of range
      */
     public int calculateInternalRotation() throws OutOfRangeException
     {
         if (this.textRotation < -90 || this.textRotation > 90)
         {
             throw new OutOfRangeException("The rotation value (" + Integer.toString(this.textRotation) + "°) is out of range. Range is form -90° to +90°");
         }
         if (this.textDirection == TextDirectionValue.vertical)
         {
             return 255;
         }
         else
         {
             if (this.textRotation >= 0)
             {
                 return this.textRotation;
             }
             else
             {
                 return (90 - this.textRotation);
             }
         }            
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
         CellXf other = (CellXf)o;
         if (this.horizontalAlign != other.getHorizontalAlign()) { return false; }
         if (this.alignment != other.getAlignment()) { return false; }
         if (this.textDirection != other.getTextDirection()) { return false; }
         if (this.textRotation != other.getTextRotation()) { return false; }
         if (this.verticalAlign != other.getVerticalAlign()) { return false; }
         else { return true; }           
    }

    /**
     * Method to compare two objects for sorting purpose
     * @param o Other object to compare with this object
     * @return -1 if the other object is bigger. 0 if both objects are equal. 1 if the other object is smaller.
     */    
     @Override
     public int compareTo(CellXf o) {
         return Integer.compare(internalID, o.getInternalID());
     }

    /**
     * Method to copy the current object to a new one
     * @return Copy of the current object without the internal ID
     */     
     public CellXf copy()
     {
         CellXf copy = new CellXf();
         copy.setHorizontalAlign(this.horizontalAlign);
         copy.setAlignment(this.alignment);
         try
         {
         copy.setTextDirection(this.textDirection);
         copy.setTextRotation(this.textRotation);
         }
         catch (Exception e)
         {
             // Should never happen. Error will be thrown earlier on setting rotation in this instance
         }
         copy.setVerticalAlign(this.verticalAlign);
         return copy;
     }    


}  
