/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2016
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.lowLevel;

import java.io.ByteArrayOutputStream;
import java.io.StringReader;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.xml.sax.InputSource;
import picoxlsx4j.Cell;
import picoxlsx4j.Column;
import picoxlsx4j.Helper;
import picoxlsx4j.Metadata;
import picoxlsx4j.Workbook;
import picoxlsx4j.Worksheet;
import picoxlsx4j.exception.IOException;
import picoxlsx4j.exception.UnknownRangeException;
import picoxlsx4j.exception.UndefinedStyleException;
import picoxlsx4j.style.Border;
import picoxlsx4j.style.CellXf;
import picoxlsx4j.style.Fill;
import picoxlsx4j.style.Font;
import picoxlsx4j.style.NumberFormat;
import picoxlsx4j.style.Style;
import picoxlsx4j.style.StyleCollection;


/**
 * Class for low level handling (XML, formatting, preparing of packing)<br>This class is only for internal use. Use the high level API (e.g. class Workbook) to manipulate data and create Excel files.
 * @author Raphael Stoeckli
 */
public class LowLevel {
        
    private Workbook workbook;
  
    /**
     * Constructor with defined workbook object
     * @param workbook Workbook to process
     */
    public LowLevel(Workbook workbook)
    {
       this.workbook = workbook;  
    }
    
    /**
     * Method to save the workbook
     * @throws IOException Thrown in case of an error
     */
    public void save() throws IOException
    {
        try
        {
        this.workbook.resolveMergedCells();
        Document doc;
        Document app = createAppPropertiesDocument();
        Document core = createCorePropertiesDocument();
        Document styles = createStyleSheetDocument();
        Document book = createWorkbookDocument();        
        String file;
        Worksheet sheet;
        Packer p = new Packer();
        Packer.Relationship rel = p.createRelationship("_rels/.rels");
        rel.addRelationshipEntry("/xl/workbook.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
        rel.addRelationshipEntry("/docProps/core.xml", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
        rel.addRelationshipEntry("/docProps/app.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
        rel = p.createRelationship("xl/_rels/workbook.xml.rels");
        for(int i = 0; i < this.workbook.getWorksheets().size(); i++)
        {
            sheet = this.workbook.getWorksheets().get(i);
            doc = createWorksheetPart(sheet);
            file = "sheet" + Integer.toString(sheet.getSheetID()) + ".xml";
            rel.addRelationshipEntry("/xl/worksheets/" + file, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
            p.addPart("xl/worksheets/" + file, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", doc);
        }
        rel.addRelationshipEntry("/xl/styles.xml", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
        p.addPart("docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml", core);
        p.addPart("docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml", app);

        p.addPart("xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", book, false);
        p.addPart("xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", styles);
        p.pack(this.workbook.getFilename());
        }
        catch (Exception e)
        {
            throw new IOException("There was an error while creating the workbook document during saving. Please see the inner exception.", e);
        }
        
        
    }
   
    /**
     * Method to create a worksheet part as XML document
     * @param worksheet worksheet object to process
     * @return Formated XML document
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    private Document createWorksheetPart(Worksheet worksheet) throws IOException
    {
        worksheet.recalculateAutoFilter();
        worksheet.recalculateColumns();
        List<List<Cell>> celldata = getSortedSheetData(worksheet);
        StringBuilder sb = new StringBuilder();
        String line;
        sb.append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
        
        if (worksheet.getSelectedCells() != null)
        {
            sb.append("<sheetViews><sheetView workbookViewId=\"0\"");
            if (this.workbook.getSelectedWorksheet() == worksheet.getSheetID() - 1)
            {
                    sb.append(" tabSelected=\"1\"");
            }
            sb.append("><selection sqref=\"");
            sb.append(worksheet.getSelectedCells().toString());
            sb.append("\" activeCell=\"");
            sb.append(worksheet.getSelectedCells().StartAddress.toString());
            sb.append("\"/></sheetView></sheetViews>");
        }
        
        sb.append("<sheetFormatPr x14ac:dyDescent=\"0.25\" defaultRowHeight=\"");
        sb.append(worksheet.getDefaultRowHeight());
        sb.append("\" baseColWidth=\"");
        sb.append(worksheet.getDefaultColumnWidth());
        sb.append("\"/>");
        String colWidths = createColsString(worksheet);
        if (Helper.isNullOrEmpty(colWidths) == false)
        {
            sb.append("<cols>");
            sb.append(colWidths);
            sb.append("</cols>");
        }
        sb.append("<sheetData>");
        for(int i = 0; i < celldata.size(); i++)
        {
            line = createRowString(celldata.get(i), worksheet);
            sb.append(line + "");
        }
        sb.append("</sheetData>");
        
        sb.append(createMergedCellsString(worksheet));
        sb.append(createSheetProtectionString(worksheet));        
        if (worksheet.getAutoFilterRange() != null)
        {
            sb.append("<autoFilter ref=\"" + worksheet.getAutoFilterRange().toString() + "\"/>");
        }
        sb.append("</worksheet>");
        
        Document doc = createXMLDocument(sb.toString());
        return doc;
    }
    
    /**
     * Method to create a style sheet as XML document
     * @return Formated XML document
     * @throws UndefinedStyleException Thrown if a style was not referenced in the style sheet
     * @throws UnknownRangeException Thrown if a referenced cell was out of range
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    private Document createStyleSheetDocument() throws IOException
    {
        StyleCollection styles = this.workbook.reorganizeStyles();
        String bordersString = createStyleBorderString(styles.getBorders());
        String fillsString = createStyleFillString(styles.getFills());
        String fontsString = createStyleFontString(styles.getFonts());
        String numberFormatsString = createStyleNumberFormatString(styles.getNumberFormats());
        int numFormatCount = getNumberFormatStringCounter(styles.getNumberFormats());
        String xfsStings = createStyleXfsString(this.workbook.getStyles());
        String mruColorString = createMruColorsString(styles.getFonts(), styles.getFills());
        StringBuilder sb = new StringBuilder();  
        
        sb.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\">");
        if (numFormatCount > 0)
        {
            sb.append("<numFmts count=\"");
            sb.append(numFormatCount);
            sb.append("\">");
            sb.append(numberFormatsString + "</numFmts>");
        }
        sb.append("<fonts x14ac:knownFonts=\"1\" count=\"");
        sb.append(styles.getFonts().size());
        sb.append("\">");
        sb.append(fontsString + "</fonts>");
        sb.append("<fills count=\"");
        sb.append(styles.getFills().size());
        sb.append("\">");
        sb.append(fillsString + "</fills>");
        sb.append("<borders count=\"");
        sb.append(styles.getBorders().size());
        sb.append("\">");
        sb.append(bordersString + "</borders>");
        sb.append("<cellXfs count=\"");
        sb.append(this.workbook.getStyles().size());
        sb.append("\">");
        sb.append(xfsStings + "</cellXfs>");
        if (this.workbook.getWorkbookMetadata() != null)
        {
            if (Helper.isNullOrEmpty(mruColorString) == false && this.workbook.getWorkbookMetadata().isUseColorMRU() == true)
            {
                sb.append("<colors>");
                sb.append(mruColorString);
                sb.append("</colors>");
            }
        }
        sb.append("</styleSheet>");
        Document doc = createXMLDocument(sb.toString());
        return doc;        
    }    
    
    /**
     * Method to create the app-properties (part of meta data) as XML document
     * @return Formated XML document
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    private Document createAppPropertiesDocument() throws IOException
    {
        StringBuilder sb = new StringBuilder();
        sb.append("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">");
        sb.append(createAppString());
        sb.append("</Properties>");
        Document doc = createXMLDocument(sb.toString());
        return doc;
    }    
    
    /**
     * Method to create the core-properties (part of meta data) as XML document
     * @return Formated XML document
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    private Document createCorePropertiesDocument() throws IOException
    {
        StringBuilder sb = new StringBuilder();
        sb.append("<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">");
        sb.append(createCorePropertiesString());
        sb.append("</cp:coreProperties>");
        Document doc = createXMLDocument(sb.toString());
        return doc;
    }    
    
    /**
     * Method to create a workbook as XML document
     * @return Formated XML document
     * @throws UnknownRangeException Thrown if a referenced cell was out of range
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    private Document createWorkbookDocument() throws IOException
    {
        if (this.workbook.getWorksheets().isEmpty())
        {
            throw new UnknownRangeException("The workbook can not be created because no worksheet was defined.");
        }
        StringBuilder sb = new StringBuilder();
        sb.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
        if (this.workbook.getSelectedWorksheet() > 0)
        {
            sb.append("<bookViews><workbookView activeTab=\"");
            sb.append(Integer.toString(this.workbook.getSelectedWorksheet()));
            sb.append("\"/></bookViews>");
        }        
        if (this.workbook.isWorkbookProtectionUsed() == true)
        {
            sb.append("<workbookProtection");
            if (this.workbook.isWindowsLockedIfProtected() == true)
            {
                sb.append(" lockWindows=\"1\"");
            }
            if (this.workbook.isStructureLockedIfProtected() == true)
            {
                sb.append(" lockStructure=\"1\"");
            }
            if (Helper.isNullOrEmpty(this.workbook.getWorkbookProtectionPassword()) == false)
            {
                sb.append("workbookPassword=\"");
                sb.append(generatePasswordHash(this.workbook.getWorkbookProtectionPassword()));
                sb.append("\"");
            }
            sb.append("/>");
        } 
        sb.append("<sheets>");
        int id;
        for (int i = 0; i < this.workbook.getWorksheets().size(); i++)
        {
            id = this.workbook.getWorksheets().get(i).getSheetID();
            sb.append("<sheet r:id=\"rId");
            sb.append(id);
            sb.append("\" sheetId=\"");
            sb.append(id);
            sb.append("\" name=\"" + LowLevel.escapeXMLAttributeChars(this.workbook.getWorksheets().get(i).getSheetName()) + "\"/>");
        }
        sb.append("</sheets>");
        sb.append("</workbook>");
        Document doc = createXMLDocument(sb.toString());
        return doc;
    }    
    
    /**
     * Method to create the XML string for the font part of the style sheet document
     * @param fontStyles List of Font objects
     * @return String with formated XML data
     */
    private String createStyleFontString(List<Font> fontStyles)
    {
        StringBuilder sb = new StringBuilder();
        Font item;
        for(int i = 0; i < fontStyles.size(); i++)
        {
            item = fontStyles.get(i);
            sb.append("<font>");
            if (item.isBold() == true) { sb.append("<b/>"); }
            if (item.isItalic() == true) { sb.append("<i/>"); }
            if (item.isUnderline() == true) { sb.append("<u/>"); }
            if (item.isDoubleUnderline() == true) { sb.append("<u val=\"double\"/>"); }
            if (item.isStrike() == true) { sb.append("<strike/>"); }
            if (item.getVerticalAlign() == Font.VerticalAlignValue.subscript) { sb.append("<vertAlign val=\"subscript\"/>"); }
            else if (item.getVerticalAlign() == Font.VerticalAlignValue.superscript) { sb.append("<vertAlign val=\"superscript\"/>"); }
            sb.append("<sz val=\"");
            sb.append(item.getSize());
            sb.append("\"/>");
            if (Helper.isNullOrEmpty(item.getColorValue()))
            {
                sb.append("<color theme=\"");
                sb.append(item.getColorTheme());
                sb.append("\"/>");
            }
            else
            {
                sb.append("<color rgb=\"" + item.getColorValue() + "\"/>");
            }
            sb.append("<name val=\"" + item.getName() + "\"/>");
            sb.append("<family val=\"" + item.getFamily() + "\"/>");
            if (item.getScheme() != Font.SchemeValue.none)
            {
                if (item.getScheme() == Font.SchemeValue.major)
                { sb.append("<scheme val=\"major\"/>"); }
                else if (item.getScheme() == Font.SchemeValue.minor)
                { sb.append("<scheme val=\"minor\"/>"); }
            }
            if (Helper.isNullOrEmpty(item.getCharset()) == false)
            {
                sb.append("<charset val=\"" + item.getCharset() + "\"/>");
            }
            sb.append("</font>");
        }
        return sb.toString();
    }    
    
    /**
     * Method to create the XML string for the border part of the style sheet document
     * @param borderStyles List of Border objects
     * @return String with formated XML data
     */
    private String createStyleBorderString(List<Border> borderStyles)
    {
        StringBuilder sb = new StringBuilder();
        Border item;
        for (int i = 0; i < borderStyles.size(); i++)
        {
            item = borderStyles.get(i);
            if (item.isDiagonalDown() == true && item.isDiagonalUp() == false) { sb.append("<border diagonalDown=\"1\">"); }
            else if (item.isDiagonalDown() == false && item.isDiagonalUp() == true) { sb.append("<border diagonalUp=\"1\">"); }
            else if (item.isDiagonalDown() == true && item.isDiagonalUp() == true) { sb.append("<border diagonalDown=\"1\" diagonalUp=\"1\">"); }
            else { sb.append("<border>"); }

            if (item.getLeftStyle() != Border.StyleValue.none)
            {
                sb.append("<left style=\"" + Border.getStyleName(item.getLeftStyle()) + "\">");
                if (Helper.isNullOrEmpty(item.getLeftColor()) == true) { sb.append("<color rgb=\"" + item.getLeftColor() + "\"/>"); }
                else { sb.append("<color auto=\"1\"/>"); }
                sb.append("</left>");
            }
            else
            {
                sb.append("<left/>");
            }
            if (item.getRightStyle() != Border.StyleValue.none)
            {
                sb.append("<right style=\"" + Border.getStyleName(item.getRightStyle()) + "\">");
                if (Helper.isNullOrEmpty(item.getRightColor()) == true) { sb.append("<color rgb=\"" + item.getRightColor() + "\"/>"); }
                else { sb.append("<color auto=\"1\"/>"); }
                sb.append("</right>");
            }
            else
            {
                sb.append("<right/>");
            }
            if (item.getTopStyle() != Border.StyleValue.none)
            {
                sb.append("<top style=\"" + Border.getStyleName(item.getTopStyle()) + "\">");
                if (Helper.isNullOrEmpty(item.getTopColor()) == true) { sb.append("<color rgb=\"" + item.getTopColor() + "\"/>"); }
                else { sb.append("<color auto=\"1\"/>"); }
                sb.append("</top>");
            }
            else
            {
                sb.append("<top/>");
            }
            if (item.getBottomStyle() != Border.StyleValue.none)
            {
                sb.append("<bottom style=\"" + Border.getStyleName(item.getBottomStyle()) + "\">");
                if (Helper.isNullOrEmpty(item.getBottomColor()) == true) { sb.append("<color rgb=\"" + item.getBottomColor() + "\"/>"); }
                else { sb.append("<color auto=\"1\"/>"); }
                sb.append("</bottom>");
            }
            else
            {
                sb.append("<bottom/>");
            }
            if (item.getDiagonalStyle() != Border.StyleValue.none)
            {
                sb.append("<diagonal style=\"" + Border.getStyleName(item.getDiagonalStyle()) + "\">");
                if (Helper.isNullOrEmpty(item.getDiagonalColor()) == true) { sb.append("<color rgb=\"" + item.getDiagonalColor() + "\"/>"); }
                else { sb.append("<color auto=\"1\"/>"); }
                sb.append("</diagonal>");
            }
            else
            {
                sb.append("<diagonal/>");
            }

            sb.append("</border>");
        }
        return sb.toString();
    }    
    
    /**
     * Method to create the XML string for the fill part of the style sheet document
     * @param fillStyles List of Fill objects
     * @return String with formated XML data
     */
    private String createStyleFillString(List<Fill> fillStyles)
    {
        StringBuilder sb = new StringBuilder();
        Fill item;
        for (int i = 0; i < fillStyles.size(); i++)
        {
            item = fillStyles.get(i);
            sb.append("<fill>");
            sb.append("<patternFill patternType=\"" + Fill.getPatternName(item.getPatternFill()) + "\"");
            if (item.getPatternFill() == Fill.PatternValue.solid)
            {
                sb.append(">");
                sb.append("<fgColor rgb=\"" + item.getForegroundColor() + "\"/>");
                sb.append("<bgColor indexed=\"");
                sb.append(item.getIndexedColor());
                sb.append("\"/>");
                sb.append("</patternFill>");
            }
            else if (item.getPatternFill() == Fill.PatternValue.mediumGray || item.getPatternFill() == Fill.PatternValue.lightGray || item.getPatternFill() == Fill.PatternValue.gray0625 || item.getPatternFill() == Fill.PatternValue.darkGray)
            {
                sb.append(">");
                sb.append("<fgColor rgb=\"" + item.getForegroundColor() + "\"/>");
                if (Helper.isNullOrEmpty(item.getBackgroundColor()) == false)
                {
                    sb.append("<bgColor rgb=\"" + item.getBackgroundColor() + "\"/>");
                }
                sb.append("</patternFill>");
            }
            else
            {
                sb.append("/>");
            }
            sb.append("</fill>");
        }
        return sb.toString();
    }  
    
    /**
     * Method to create the XML string for the color-MRU part of the style sheet document (recent colors)
     * @param fonts List of Font objects
     * @param fills List of Fill objects
     * @return String with formated XML data
     */
    private String createMruColorsString(List<Font> fonts, List<Fill> fills)
    {
        StringBuilder sb = new StringBuilder();
        List<String> tempColors = new ArrayList<>();
        Font item;
        for (int i = 0; i < fonts.size(); i++)
        {
            item = fonts.get(i);
            if (Helper.isNullOrEmpty(item.getColorValue()) == true) { continue; }
            if (item.getColorValue().equals(Fill.DEFAULTCOLOR)) { continue; }
            if (tempColors.contains(item.getColorValue()) == false) { tempColors.add(item.getColorValue()); }
        }
        Fill item2;
        for (int i = 0; i < fills.size(); i++)
        {
            item2 = fills.get(i);
            if (Helper.isNullOrEmpty(item2.getBackgroundColor()) == false)
            {
                if (item2.getBackgroundColor().equals(Fill.DEFAULTCOLOR) == false)
                {
                    if (tempColors.contains(item2.getBackgroundColor()) == false) { tempColors.add(item2.getBackgroundColor()); }
                }
            }
            if (Helper.isNullOrEmpty(item2.getForegroundColor()) == false)
            {
                if (item2.getForegroundColor().equals(Fill.DEFAULTCOLOR) == false)
                {
                    if (tempColors.contains(item2.getForegroundColor()) == false) { tempColors.add(item2.getForegroundColor()); }
                }
            }
        }
        if (tempColors.size() > 0)
        {
            sb.append("<mruColors>");
            for(int i = 0; i < tempColors.size(); i++)
            {
                sb.append("<color rgb=\"" + tempColors.get(i) + "\"/>");
            }
            sb.append("</mruColors>");
            return sb.toString();
        }
        else
        {
            return "";
        }
    }    
    
    /**
     * Method to create the XML string for the XF part of the style sheet document
     * @param styles List of Style objects
     * @return String with formated XML data
     * @throws UnknownRangeException Thrown if a referenced cell was out of range
     */
    private String createStyleXfsString(List<Style> styles)
    {
        StringBuilder sb = new StringBuilder();
        StringBuilder sb2 = null;
        String alignmentString, protectionString;
        int formatNumber, textRotation;
        Style item;
        for (int i = 0; i < styles.size(); i++)
        {
            item = styles.get(i);
            textRotation = item.getCurrentCellXf().calculateInternalRotation();
            alignmentString = "";
            protectionString = "";
            if (item.getCurrentCellXf().getHorizontalAlign() != CellXf.HorizontalAlignValue.none || item.getCurrentCellXf().getVerticalAlign() != CellXf.VerticallAlignValue.none || item.getCurrentCellXf().getAlignment() != CellXf.TextBreakValue.none || textRotation != 0)
            {
                sb2 = new StringBuilder();
                sb2.append("<alignment");
                if (item.getCurrentCellXf().getHorizontalAlign() != CellXf.HorizontalAlignValue.none)
                {
                    sb2.append(" horizontal=\"");
                    if (item.getCurrentCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.center) { sb2.append("center"); }
                    else if (item.getCurrentCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.right) { sb2.append("right"); }
                    else if (item.getCurrentCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.centerContinuous) { sb2.append("centerContinuous"); }
                    else if (item.getCurrentCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.distributed) { sb2.append("distributed"); }
                    else if (item.getCurrentCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.fill) { sb2.append("fill"); }
                    else if (item.getCurrentCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.general) { sb2.append("general"); }
                    else if (item.getCurrentCellXf().getHorizontalAlign() == CellXf.HorizontalAlignValue.justify) { sb2.append("justify"); }
                    else { sb2.append("left"); }
                    sb2.append("\"");
                }
                if (item.getCurrentCellXf().getVerticalAlign() != CellXf.VerticallAlignValue.none)
                {
                    sb2.append(" vertical=\"");
                    if (item.getCurrentCellXf().getVerticalAlign() == CellXf.VerticallAlignValue.center) { sb2.append("center"); }
                    else if (item.getCurrentCellXf().getVerticalAlign() == CellXf.VerticallAlignValue.distributed) { sb2.append("distributed"); }
                    else if (item.getCurrentCellXf().getVerticalAlign() == CellXf.VerticallAlignValue.justify) { sb2.append("justify"); }
                    else if (item.getCurrentCellXf().getVerticalAlign() == CellXf.VerticallAlignValue.top) { sb2.append("top"); }
                    else { sb2.append("bottom"); }
                    sb2.append("\"");
                }

                if (item.getCurrentCellXf().getAlignment() != CellXf.TextBreakValue.none)
                {
                    if (item.getCurrentCellXf().getAlignment() == CellXf.TextBreakValue.shrinkToFit) { sb2.append(" shrinkToFit=\"1"); }
                    else { sb2.append(" wrapText=\"1"); }
                    sb2.append("\"");
                }
                if (textRotation != 0)
                {
                    sb2.append(" textRotation=\"");
                    sb2.append(textRotation);
                    sb2.append("\"");
                }
                sb2.append("/>"); // </xf>
                alignmentString = sb2.toString();
            }
            
            if (item.getCurrentCellXf().isHidden() == true || item.getCurrentCellXf().isLocked() == true)
            {
                if (item.getCurrentCellXf().isHidden() == true && item.getCurrentCellXf().isLocked() == true)
                {
                    protectionString = "<protection locked=\"1\" hidden=\"1\"/>";
                }
                else if (item.getCurrentCellXf().isHidden() == true && item.getCurrentCellXf().isLocked() == false)
                {
                    protectionString = "<protection hidden=\"1\" locked=\"0\"/>";
                }
                else
                {
                    protectionString = "<protection hidden=\"0\" locked=\"1\"/>";
                }
            }
            
            sb.append("<xf numFmtId=\"");
            if (item.getCurrentNumberFormat().isCustomFormat() == true)
            {
                sb.append(item.getCurrentNumberFormat().getCustomFormatID());
            }
            else
            {
                formatNumber = item.getCurrentNumberFormat().getNumber().getNumVal();
                sb.append(formatNumber);
            }
            sb.append("\" borderId=\"");
            sb.append(item.getCurrentBorder().getInternalID());
            sb.append("\" fillId=\"");
            sb.append(item.getCurrentFill().getInternalID());
            sb.append("\" fontId=\"");
            sb.append(item.getCurrentFont().getInternalID());
            if (item.getCurrentFont().isDefaultFont() == false)
            {
                sb.append("\" applyFont=\"1");
            }
            if (item.getCurrentFill().getPatternFill() != Fill.PatternValue.none)
            {
                sb.append("\" applyFill=\"1");
            }
            if (item.getCurrentBorder().isEmpty() == false)
            {
                sb.append("\" applyBorder=\"1");
            }
            if (alignmentString.isEmpty() == false || item.getCurrentCellXf().isForceApplyAlignment() == true)
            {
                sb.append("\" applyAlignment=\"1");
            }
            if (protectionString.isEmpty() == false)
            {
                sb.append("\" applyProtection=\"1");
            }            
            if (item.getCurrentNumberFormat().getNumber() != NumberFormat.FormatNumber.none)
            {
                sb.append("\" applyNumberFormat=\"1\"");
            }
            else
            {
                sb.append("\""); 
            }
            if (alignmentString.isEmpty() == false  || protectionString.isEmpty() == false)
            {
                sb.append(">");
                sb.append(alignmentString);
                sb.append(protectionString);
                sb.append("</xf>");
            }
            else
            {
                sb.append("/>");
            }
        }
        return sb.toString();
    }    
    
    /**
     * Method to create the XML string for the number format part of the style sheet document 
     * @param numberFormatStyles List of NumberFormat objects
     * @return String with formated XML data
     */
    private String createStyleNumberFormatString(List<NumberFormat> numberFormatStyles)
    {

        StringBuilder sb = new StringBuilder();
        NumberFormat item;
        for (int i = 0; i < numberFormatStyles.size(); i++)
        {
            item = numberFormatStyles.get(i);
            if (item.isCustomFormat() == true)
            {
                sb.append("<numFmt formatCode=\"" + item.getCustomFormatCode() + "\" numFmtId=\"");
                sb.append(item.getCustomFormatID());
                sb.append("\"/>");
            }
        }
        return sb.toString();
    }
    
    /**
     * Gets the number of custom number formats
     * @param numberFormatStyles List of NumberFormat objects
     * @return Number of custom number formats to apply in the style document
     */
    private int getNumberFormatStringCounter(List<NumberFormat> numberFormatStyles)
    {
        int counter = 0;
        NumberFormat item;
        for (int i = 0; i < numberFormatStyles.size(); i++)
        {
            item = numberFormatStyles.get(i);
            if (item.isCustomFormat() == true)
            {
                counter++;
            }
        }
        return counter;
    }
    
    /**
     * Method to create the columns as XML string. This is used to define the width of columns
     * @param worksheet Worksheet to process
     * @return String with formated XML data
     */
    private String createColsString(Worksheet worksheet)
    {
        if (worksheet.getColumns().size() > 0)
        {
            String col;
            String hidden = "";
            StringBuilder sb = new StringBuilder();
            
            //Iterator itr = worksheet.getColumnWidths().entrySet().iterator();
            //Map.Entry<Integer, Float> width;
            //while (itr.hasNext())
            for (Map.Entry<Integer, Column> column : worksheet.getColumns().entrySet())
            {  
                //width = (Map.Entry<Integer, Float>)itr.next();
                if (column.getValue().getWidth() == worksheet.getDefaultColumnWidth() && column.getValue().isHidden() == false) { continue; }
                if (worksheet.getColumns().containsKey(column.getKey()))
                {
                    if (worksheet.getColumns().get(column.getKey()).isHidden() == true)
                    {
                        hidden = " hidden=\"1\"";
                    }
                }
                col = Integer.toString(column.getKey() + 1); // Add 1 for Address
                sb.append("<col customWidth=\"1\" width=\"" + Float.toString(column.getValue().getWidth()) + "\" max=\"" + col + "\" min=\"" + col + "\"" + hidden + "/>");
            }            
            String value = sb.toString();
            if (value.length() > 0)
            {
                return value;
            }
            else
            {
                return "";
            }
        }
        else
        {
            return "";
        }
    }    
        
    /**
     * Method to create a row string
     * @param columnFields List of cells
     * @param worksheet Worksheet to process
     * @return Formated row string
     */
    private String createRowString(List<Cell> columnFields, Worksheet worksheet)
    {
        int rowNumber = columnFields.get(0).getRowAddress();
        String heigth = "";
        String hidden = "";
        if (worksheet.getRowHeights().containsKey(rowNumber))
        {
            if (worksheet.getRowHeights().get(rowNumber) != worksheet.getDefaultRowHeight())
            {
                heigth = " x14ac:dyDescent=\"0.25\" customHeight=\"1\" ht=\"" + Float.toString(worksheet.getRowHeights().get(rowNumber)) + "\"";
            }
        }
        if (worksheet.getHiddenRows().containsKey(rowNumber))
        {
            if (worksheet.getHiddenRows().get(rowNumber) == true)
            {
                hidden = " hidden=\"1\"";
            }
        }        
        StringBuilder sb = new StringBuilder();
        if (columnFields.size() > 0)
        {
            sb.append("<row r=\"");
            sb.append((rowNumber + 1));
            sb.append("\"" + heigth + hidden + ">");
        }
        else
        {
            sb.append("<row" + heigth + ">");
        }
        String typeAttribute;
        String sValue = "";
        String tValue = "";
        String value = "";
        boolean bVal;

        Date dVal;
        int col = 0;
        Cell item;
        for (int i = 0; i < columnFields.size(); i++)
        {
            item = columnFields.get(i);
            tValue = " ";
            if (item.getCellStyle() != null)
            {
                sValue = " s=\"" + Integer.toString(item.getCellStyle().getInternalID()) + "\" ";
            }
            else
            {
                sValue = "";
            }
            item.resolveCellType(); // Recalculate the type (for handling DEFAULT)
            if (item.getFieldType() == Cell.CellType.BOOL)
            {
                typeAttribute = "b";
                tValue = " t=\"" + typeAttribute + "\" ";
                bVal = (boolean)item.getValue();
                if (bVal == true) { value = "1"; }
                else { value = "0"; }

            }
            // Number casting
            else if (item.getFieldType() == Cell.CellType.NUMBER)
            {
                typeAttribute = "n";
                tValue = " t=\"" + typeAttribute + "\" ";
                Class t = item.getValue().getClass();


                if (t.equals(Integer.class))
                {
                    value = Integer.toString((int)item.getValue());
                }
                else if(t.equals(Double.class))
                {
                    value = Double.toString((double)item.getValue());

                }
                else if (t.equals(Float.class))
                {
                    value = Float.toString((float)item.getValue());
                }

            }
            // Date parsing
            else if (item.getFieldType() == Cell.CellType.DATE)
            {
                typeAttribute = "d";
                dVal = (Date)item.getValue();
                value = Double.toString(Helper.getOADate(dVal));
            }
            // String parsing
            else
            {
                typeAttribute = "str";
                tValue = " t=\"" + typeAttribute + "\" ";      
                if (item.getValue() == null)
                {
                    value = "";
                }
                else
                {
                    value = item.getValue().toString();
                }                
            }
            if (item.getFieldType() != Cell.CellType.EMPTY)
            {
                sb.append("<c" + tValue + "r=\"" + item.getCellAddress() + "\"" + sValue + ">");
                if (item.getFieldType() == Cell.CellType.FORMULA)
                {
                    sb.append("<f>" + LowLevel.escapeXMLChars(item.getValue().toString()) + "</f>");
                }
                else
                {
                    sb.append("<v>" + LowLevel.escapeXMLChars(value) + "</v>");
                }
                sb.append("</c>");
            }
            else // Empty cell
            {
                sb.append("<c" + tValue + "r=\"" + item.getCellAddress() + "\"" + sValue + "/>");
            }
            col++;
        }
        sb.append("</row>");
        return sb.toString();
    }
    
    /**
     * Method to create the merged cells string of the passed worksheet
     * @param sheet Worksheet to process
     * @return Formated string with merged cell ranges
     */
    private String createMergedCellsString(Worksheet sheet)
   {
       if (sheet.getMergedCells().size() < 1)
       {
           return "";
       }
            Iterator itr;
            Map.Entry<String, Cell.Range> range;
            StringBuilder sb = new StringBuilder();
            sb.append("<mergeCells count=\"" + Integer.toString(sheet.getMergedCells().size()) + "\">");
            itr = sheet.getMergedCells().entrySet().iterator();
            while (itr.hasNext())
            {
            range = (Map.Entry<String, Cell.Range>)itr.next();
            sb.append("<mergeCell ref=\"" + range.getValue().toString() + "\"/>");
            }
       sb.append("</mergeCells>");
       return sb.toString();
   } 
   
    /**
     * Method to create the protection string of the passed worksheet
     * @param sheet Worksheet to process
     * @return Formated string with protection statement of the worksheet
     */
    private String createSheetProtectionString(Worksheet sheet)
    {
        if (sheet.isUseSheetProtection() == false)
        {
            return "";
        }
        HashMap<Worksheet.SheetProtectionValue, Integer> actualLockingValues = new HashMap<Worksheet.SheetProtectionValue,Integer>();
        if (sheet.getSheetProtectionValues().size() == 0)
        {
            actualLockingValues.put(Worksheet.SheetProtectionValue.selectLockedCells, 1);
            actualLockingValues.put(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
        }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.objects) == false)
        {
            actualLockingValues.put(Worksheet.SheetProtectionValue.objects, 1);
        }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.scenarios) == false)
        {
            actualLockingValues.put(Worksheet.SheetProtectionValue.scenarios, 1);
        }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.selectLockedCells) == false )
        {
            if (actualLockingValues.containsKey(Worksheet.SheetProtectionValue.selectLockedCells) == false)
            {
                actualLockingValues.put(Worksheet.SheetProtectionValue.selectLockedCells, 1);
            }            
        }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.selectUnlockedCells) == false || sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.selectLockedCells) == false)
        {
            if (actualLockingValues.containsKey(Worksheet.SheetProtectionValue.selectUnlockedCells) == false)
            {
                actualLockingValues.put(Worksheet.SheetProtectionValue.selectUnlockedCells, 1);
            }            
        }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.formatCells)) { actualLockingValues.put(Worksheet.SheetProtectionValue.formatCells, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.formatColumns)) { actualLockingValues.put(Worksheet.SheetProtectionValue.formatColumns, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.formatRows)) { actualLockingValues.put(Worksheet.SheetProtectionValue.formatRows, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.insertColumns)) { actualLockingValues.put(Worksheet.SheetProtectionValue.insertColumns, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.insertRows)) { actualLockingValues.put(Worksheet.SheetProtectionValue.insertRows, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.insertHyperlinks)) { actualLockingValues.put(Worksheet.SheetProtectionValue.insertHyperlinks, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.deleteColumns)) { actualLockingValues.put(Worksheet.SheetProtectionValue.deleteColumns, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.deleteRows)) { actualLockingValues.put(Worksheet.SheetProtectionValue.deleteRows, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.sort)) { actualLockingValues.put(Worksheet.SheetProtectionValue.sort, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.autoFilter)) { actualLockingValues.put(Worksheet.SheetProtectionValue.autoFilter, 0); }
        if (sheet.getSheetProtectionValues().contains(Worksheet.SheetProtectionValue.pivotTables)) { actualLockingValues.put(Worksheet.SheetProtectionValue.pivotTables, 0); }
        StringBuilder sb = new StringBuilder();
        sb.append("<sheetProtection");
        String temp;
        Iterator itr;
        Map.Entry<Worksheet.SheetProtectionValue, Integer> item;        
        itr = actualLockingValues.entrySet().iterator();
        while (itr.hasNext())
        {
            item = (Map.Entry<Worksheet.SheetProtectionValue, Integer>)itr.next();
            temp = item.getKey().name();// Note! If the enum names differs from the OOXML definitions, this method will cause invalid OOXML entries
         }
            if (Helper.isNullOrEmpty(sheet.getSheetProtectionPassword()) == false)
            {
                String hash = generatePasswordHash(sheet.getSheetProtectionPassword());
                sb.append(" password=\"" + hash + "\"");
            }        
        sb.append(" sheet=\"1\"/>");
       return sb.toString();
    }    
    
    /**
     * Method to create the XML string for the app-properties document
     * @return String with formated XML data
     */
    private String createAppString()
    {
        if (this.workbook.getWorkbookMetadata() == null) { return ""; }
        Metadata md = this.workbook.getWorkbookMetadata();
        StringBuilder sb = new StringBuilder();
        appendXMLtag(sb, "0", "TotalTime", null);
        appendXMLtag(sb, md.getApplication(), "Application", null);
        appendXMLtag(sb, "0", "DocSecurity", null);
        appendXMLtag(sb, "false", "ScaleCrop", null);
        appendXMLtag(sb, md.getManager(), "Manager", null);
        appendXMLtag(sb, md.getCompany(), "Company", null);
        appendXMLtag(sb, "false", "LinksUpToDate", null);
        appendXMLtag(sb, "false", "SharedDoc", null);
        appendXMLtag(sb, md.getHyperlinkBase(), "HyperlinkBase", null);
        appendXMLtag(sb, "false", "HyperlinksChanged", null);
        appendXMLtag(sb, md.getApplicationVersion(), "AppVersion", null);
        return sb.toString();
    }    

    /**
     * Method to create the XML string for the core-properties document
     * @return String with formated XML data
     */
    private String createCorePropertiesString()
    {
        if (this.workbook.getWorkbookMetadata() == null) { return ""; }
        Metadata md = this.workbook.getWorkbookMetadata();
        StringBuilder sb = new StringBuilder();
        appendXMLtag(sb, md.getTitle(), "title", "dc");
        appendXMLtag(sb, md.getSubject(), "subject", "dc");
        appendXMLtag(sb, md.getCreator(), "creator", "dc");
        appendXMLtag(sb, md.getCreator(), "lastModifiedBy", "cp");
        appendXMLtag(sb, md.getKeywords(), "keywords", "cp");
        appendXMLtag(sb, md.getDescription(), "description", "dc");
        
        DateFormat  df = new SimpleDateFormat("yyyy-MM-dd'T'hh:mm:ss'Z'");
        Date now = Calendar.getInstance().getTime();
        String time = df.format(now);
        
        //string time = DateTime.Now.ToString("yyyy-MM-ddThh:mm:ssZ");
        sb.append("<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + time + "</dcterms:created>");
        sb.append("<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + time + "</dcterms:modified>");

        appendXMLtag(sb, md.getCategory(), "category", "cp");
        appendXMLtag(sb, md.getContentStatus(), "contentStatus", "cp");

        return sb.toString();
    }    
    
    /**
     * Method to sort the cells of a worksheet as preparation for the XML document
     * @param sheet Worksheet to process 
     * @return Two dimensional array of Cell object
     */
    private List<List<Cell>> getSortedSheetData(Worksheet sheet)
     {
         List<Cell> temp = new ArrayList<>();
        Map.Entry entry;
        Iterator itr = sheet.getCells().entrySet().iterator();
        while (itr.hasNext())
        {  
            entry = (Map.Entry)itr.next();
            temp.add((Cell)entry.getValue());
        }
         Collections.sort(temp);             
         List<Cell> line = new ArrayList<>();
         List<List<Cell>> output = new ArrayList<>();
         if (temp.size() > 0)
         {
             int rowNumber = temp.get(0).getRowAddress();
             for (int i = 0; i < temp.size(); i++)
             {
                 if (temp.get(i).getRowAddress() != rowNumber)
                 {
                     output.add(line);
                     line = new ArrayList<>();
                     rowNumber = temp.get(i).getRowAddress();
                 }
                 line.add(temp.get(i));
             }
             if (line.size() > 0)
             {
                 output.add(line);
             }
         }
         return output;
     }    
    
    /**
     * Method to escape XML characters between two XML tags
     * @param input Input string to process
     * @return Escaped string
     */
    public static String escapeXMLChars(String input)
    {
        input = input.replace("<", "&lt;");
        input = input.replace(">", "&gt;");
        return input;
    }
    
    /**
     * Method to escape XML characters in an XML attribute
     * @param input Input string to process
     * @return Escaped string
     */
    public static String escapeXMLAttributeChars(String input)
    {
        input = input.replace("\"", "&quot;");
        return input;
    }    
    
    /**
     *  Method to append a simple XML tag with an enclosed value to the passed StringBuilder
     * @param sb StringBuilder to append
     * @param value Value of the XML element
     * @param tagName Tag name of the XML element
     * @param nameSpace Optional XML name space. Can be empty or null
     * @return Returns false if no tag was appended, because the value or tag name was null or empty
     */
    private boolean appendXMLtag(StringBuilder sb, String value, String tagName, String nameSpace)
    {
        if (Helper.isNullOrEmpty(value)) { return false; }
        if (sb == null || Helper.isNullOrEmpty(tagName)) { return false; }
        boolean hasNoNs = Helper.isNullOrEmpty(nameSpace);
        sb.append('<');
        if (hasNoNs == false)
        {
            sb.append(nameSpace);
            sb.append(':');
        }
        sb.append(tagName + ">");
        sb.append(escapeXMLChars(value));
        sb.append("</");
        if (hasNoNs == false)
        {
            sb.append(nameSpace);
            sb.append(':');
        }
        sb.append(tagName);
        sb.append(">");
        return true;
    }    
    
    /**
     * Creates a XML document from a string
     * @param rawInput String to process
     * @return Formated XML document
     * @throws IOException Thrown in case of an error while creating the XML document
     */
    public static Document createXMLDocument(String rawInput) throws IOException
    {
        try
        {
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();   
        DocumentBuilder docBuilder = factory.newDocumentBuilder();
        InputSource input = new InputSource(new StringReader( rawInput ));
        input.setEncoding("UTF-8");
        Document doc = docBuilder.parse( input );
        doc.setXmlVersion("1.0");
        doc.setXmlStandalone(true);
        return doc;
        }
        catch(Exception e)
        {
            throw new IOException("There was an error while creating the XML document. Please see the inner exception.", e);
        }
    }
    
    /**
     * Method to convert an XML document to an byte array
     * @param document Document to process
     * @return array of bytes (UTF-8)
     * @throws IOException Thrown if the document could not be converted to a byte stream
     */
    public static byte[] createBytesFromDocument(Document document) throws IOException
    {
        try
        {
        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
        ByteArrayOutputStream bs = new ByteArrayOutputStream();
        Result output = new StreamResult(bs);
        Source input = new DOMSource(document);

        transformer.transform(input, output);
        bs.flush();
        byte[] bytes = bs.toByteArray();
        bs.close();
        return bytes;
        }
        catch(Exception e)
        {
            throw new IOException("There was an error while creating the byte stream. Please see the inner exception.", e);
        }  
    }
    
    /**
     * Method to generate an Excel internal password hash to protect workbooks or worksheets<br>
     * This method is derived from the c++ implementation by Kohei Yoshida (<a href="http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/">http://kohei.us/2008/01/18/excel-sheet-protection-password-hash/</a>)<br>
     * <b>WARNING!</b> Do not use this method to encrypt 'real' passwords or data outside from PicoXLSX4j. This is only a minor security feature. Use a proper cryptography method instead.
     * @param password Password as plain text
     * @return Encoded password
     */
    public static String generatePasswordHash(String password)
    {
        if (Helper.isNullOrEmpty(password)) { return ""; }
        int PasswordLength = password.length();
        int passwordHash = 0;
        char character;
        for(int i = PasswordLength; i > 0; i--)
        {
            character = password.charAt(i - 1);
            passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
            passwordHash ^= character;
        }
        passwordHash = ((passwordHash >> 14) & 0x01) | ((passwordHash << 1) & 0x7fff);
        passwordHash ^= (0x8000 | ('N' << 8) | 'K');
        passwordHash ^= PasswordLength;
        return Integer.toHexString(passwordHash).toUpperCase();
    }    
    
}
