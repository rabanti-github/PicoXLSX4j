/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.style;

/**
 * Factory class with the most important predefined styles
 * @author Raphael stoeckli
 */
public final class BasicStyles {
    /**
     * Enum with style selection
     */
    private enum StyleEnum
    {
        /**
        * Format text bold
        */
        bold,
        /**
        * Format text italic
        */
        italic,
        /**
        * Format text bold and italic
        */
        boldItalic,
        /**
        * Format text with an underline
        */
        underline,
        /**
        * Format text with a double underline
        */
        doubleUnderlien,
        /**
        * Format text with a strike-through
        */
        strike,
        /**
        * Format number as date
        */
        dateFormat,
        /**
        * Rounds number as an integer
        */
        roundFormat,
        /**
        * Format cell with a thin border
        */
        borderFrame,
        /**
        * Format cell with a thin border and a thick bottom line as header cell
        */
        borderFrameHeader,
        /**
        * Special pattern fill style for compatibility purpose 
        */
        dottedFill_0_125
    }    
     private static Style bold, italic, boldItalic, underline, doubleUnderline, strike, dateFormat, roundFormat, borderFrame, borderFrameHeader, dottedFill_0_125;
     
    /**
    * Gets the bold style
    * @returns Style object
    */
    public static Style Bold()
    { return getStyle(StyleEnum.bold);}
    /**
    * Gets the italic style
    * @returns Style object
    */
    public static Style Italic()
    { return getStyle(StyleEnum.italic);}
    /**
    * Gets the bold and italic style
    * @returns Style object
    */
    public static Style BoldItalic()
    { return getStyle(StyleEnum.boldItalic);}
    /**
    * Gets the underline style
    * @returns Style object
    */
    public static Style Underline()
    { return getStyle(StyleEnum.underline);}
    /**
    * Gets the double underline style
    * @returns Style object
    */
    public static Style DoubleUnderline()
    { return getStyle(StyleEnum.doubleUnderlien);}
    /**
    * Gets the strike style
    * @returns Style object
    */
    public static Style Strike()
    { return getStyle(StyleEnum.strike);}
    /**
    * Gets the date format style
    * @returns Style object
    */
    public static Style DateFormat()
    { return getStyle(StyleEnum.dateFormat);}
    /**
    * Gets the round format style
    * @returns Style object
    */
    public static Style RoundFormat()
    { return getStyle(StyleEnum.roundFormat);}
    /**
    * Gets the border frame style
    * @returns Style object
    */
    public static Style BorderFrame()
    { return getStyle(StyleEnum.borderFrame);}
    /**
    * Gets the border style for header cells
    * @returns Style object
    */
    public static Style BorderFrameHeader()
    { return getStyle(StyleEnum.borderFrameHeader);}
    /**
    * Gets the special pattern fill style (for compatibility)
    * @returns Style object
    */
    public static Style DottedFill_0_125()
    { return getStyle(StyleEnum.dottedFill_0_125);}       
     
    /**
     * Method to maintain the styles and to create singleton instances
     * @param value Enum value to maintain
     * @return The style according to the passed enum value
     */
    private static Style getStyle(StyleEnum value)
    {
        Style s = null;
        switch (value)
        {
            case bold:
                if (bold == null)
                {
                    bold = new Style();
                    bold.getCurrentFont().setBold(true);
                }
                s = bold;
                break;
            case italic:
                if (italic == null)
                {
                    italic = new Style();
                    italic.getCurrentFont().setItalic(true);
                }
                s = italic;
                break;
            case boldItalic:
                if (boldItalic == null)
                {
                    boldItalic = new Style();
                    boldItalic.getCurrentFont().setItalic(true);
                    boldItalic.getCurrentFont().setBold(true);
                }
                s = boldItalic;
                break;
            case underline:
                if (underline == null)
                {
                    underline = new Style();
                    underline.getCurrentFont().setUnderline(true);
                }
                s = underline;
                break;
            case doubleUnderlien:
                if (doubleUnderline == null)
                {
                    doubleUnderline = new Style();
                    doubleUnderline.getCurrentFont().setDoubleUnderline(true);
                }
                s = doubleUnderline;
                break;
            case strike:
                if (strike == null)
                {
                    strike = new Style();
                    strike.getCurrentFont().setStrike(true);
                }
                s = strike;
                break;
            case dateFormat:
                if (dateFormat == null)
                {
                    dateFormat = new Style();
                    dateFormat.getCurrentNumberFormat().setNumber(NumberFormat.FormatNumber.format_14);
                }
                s = dateFormat;
                break;
            case roundFormat:
                if (roundFormat == null)
                {
                    roundFormat = new Style();
                    roundFormat.getCurrentNumberFormat().setNumber(NumberFormat.FormatNumber.format_1);
                }
                s = roundFormat;
                break;
            case borderFrame:
                if (borderFrame == null)
                {
                    borderFrame = new Style();
                    borderFrame.getCurrentBorder().setTopStyle(Border.StyleValue.thin);
                    borderFrame.getCurrentBorder().setBottomStyle(Border.StyleValue.thin);
                    borderFrame.getCurrentBorder().setLeftStyle(Border.StyleValue.thin);
                    borderFrame.getCurrentBorder().setRightStyle(Border.StyleValue.thin);
                }
                s = borderFrame;
                break;
            case borderFrameHeader:
                if (borderFrameHeader == null)
                {
                    borderFrameHeader = new Style();
                    borderFrameHeader.getCurrentBorder().setTopStyle(Border.StyleValue.thin);
                    borderFrameHeader.getCurrentBorder().setBottomStyle(Border.StyleValue.medium);
                    borderFrameHeader.getCurrentBorder().setLeftStyle(Border.StyleValue.thin);
                    borderFrameHeader.getCurrentBorder().setRightStyle(Border.StyleValue.thin);
                    borderFrameHeader.getCurrentFont().setBold(true);
                }
                s = borderFrameHeader;
                break;
            case dottedFill_0_125:
                if (dottedFill_0_125 == null)
                {
                    dottedFill_0_125 = new Style();
                    dottedFill_0_125.getCurrentFill().setPatternFill(Fill.PatternValue.gray125);
                }
                s = dottedFill_0_125;
                break;
            default:
                break;
        }
        return s;
    }     
     
     
}
