/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.demo;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import picoxlsx4j.Cell;
import picoxlsx4j.Workbook;
import picoxlsx4j.Worksheet;
import picoxlsx4j.style.BasicStyles;
import picoxlsx4j.style.CellXf;
import picoxlsx4j.style.Fill;
import picoxlsx4j.style.Style;

/**
 * Demo Program for PicoXLSX4j
 * @author Raphael Stoeckli
 */
public class PicoXLSX4j {

    /**
     * Method to run all demos 
     * @param args the command line arguments (not used)
     */
    public static void main(String[] args)  {
       
        basicDemo();
        demo1();
        demo2();
        demo3();
        demo4();
        demo5();
        demo6();
    }
    
        /**
         * This is a very basic demo (adding three values and save the workbook)
         */
        private static void basicDemo()
        {
            try
            {
            Workbook workbook = new Workbook("basic.xlsx", "Sheet1");           // Create new workbook
            workbook.getCurrentWorksheet().addNextCell("Test");                 // Add cell A1
            workbook.getCurrentWorksheet().addNextCell("Test2");                // Add cell B1
            workbook.getCurrentWorksheet().addNextCell("Test3");                // Add cell C1
            workbook.save();
            }
            catch(Exception e)
            {
                System.out.println(e.getMessage());
            }
        }
 
        /**
         * This method shows the usage of AddNextCell with several data types and formulas
         */
        private static void demo1()
        {
            try
            {
            Workbook workbook = new Workbook("test1.xlsx", "Sheet1");           // Create new workbook
            workbook.getCurrentWorksheet().addNextCell("Test");                 // Add cell A1
            workbook.getCurrentWorksheet().addNextCell(123);                    // Add cell B1
            workbook.getCurrentWorksheet().addNextCell(true);                   // Add cell C1
            workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 2
            workbook.getCurrentWorksheet().addNextCell(123.456d);               // Add cell A2
            workbook.getCurrentWorksheet().addNextCell(123.789f);               // Add cell B2
            workbook.getCurrentWorksheet().addNextCell(new Date());             // Add cell C2
            workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 3
            workbook.getCurrentWorksheet().addNextCellFormula("B1*22");         // Add cell A3 as formula (B1 times 22)
            workbook.getCurrentWorksheet().addNextCellFormula("ROUNDDOWN(A2,1)"); // Add cell B3 as formula (Floor A2 with one decimal place)
            workbook.getCurrentWorksheet().addNextCellFormula("PI()");          // Add cell C3 as formula (Pi = 3.14.... )
            workbook.save();                                                    // Save the workbook
            }
            catch(Exception e)
            {
                System.out.println(e.getMessage());
            }
        }            
     
        /**
         * This demo shows the usage of several data types, the method AddCell, more than one worksheet and the SaveAs method
         */
        private static void demo2()
        {
            try
            {            
            Workbook workbook = new Workbook(false);                            // Create new workbook
            workbook.addWorksheet("Sheet1");                                    // Add a new Worksheet and set it as current sheet
            workbook.getCurrentWorksheet().addNextCell("月曜日");                // Add cell A1 (Unicode)
            workbook.getCurrentWorksheet().addNextCell(-987);                   // Add cell B1
            workbook.getCurrentWorksheet().addNextCell(false);                  // Add cell C1
            workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 2
            workbook.getCurrentWorksheet().addNextCell(-123.456d);              // Add cell A2
            workbook.getCurrentWorksheet().addNextCell(-123.789f);              // Add cell B2
            workbook.getCurrentWorksheet().addNextCell(new Date());             // Add cell C3
            workbook.addWorksheet("Sheet2");                                    // Add a new Worksheet and set it as current sheet
            workbook.getCurrentWorksheet().addCell("ABC", "A1");                // Add cell A1
            workbook.getCurrentWorksheet().addCell(779, 2, 1);                  // Add cell C2 (zero based addresses: column 2=C, row 1=2)
            workbook.getCurrentWorksheet().addCell(false, 3, 2);                // Add cell D3 (zero based addresses: column 3=D, row 2=3)
            workbook.getCurrentWorksheet().addNextCell(0);                      // Add cell E3 (direction: column to column)
            List<String> values = new ArrayList<>();                            // Create a List of values
            values.add("V1");
            values.add("V2");
            values.add("V3");
            workbook.getCurrentWorksheet().addStringCellRange(values, "A4:C4"); // Add a cell range to A4 - C4
            workbook.saveAs("test2j.xlsx");                                     // Save the workbook
            }
            catch(Exception e)
            {
                System.out.println(e.getMessage());
            }            
        }
        
        /**
         * This demo shows the usage of flipped direction when using AddnextCell
         */
        private static void demo3()
        {
            try
            { 
            Workbook workbook = new Workbook("test3.xlsx", "Sheet1");           // Create new workbook
            workbook.getCurrentWorksheet().setCurrentCellDirection(Worksheet.CellDirection.RowToRow);  // Change the cell direction
            workbook.getCurrentWorksheet().addNextCell(1);                      // Add cell A1
            workbook.getCurrentWorksheet().addNextCell(2);                      // Add cell A2
            workbook.getCurrentWorksheet().addNextCell(3);                      // Add cell A3
            workbook.getCurrentWorksheet().addNextCell(4);                      // Add cell A4
            workbook.getCurrentWorksheet().goToNextColumn();                    // Go to Column B
            workbook.getCurrentWorksheet().addNextCell("A");                    // Add cell B1
            workbook.getCurrentWorksheet().addNextCell("B");                    // Add cell B2
            workbook.getCurrentWorksheet().addNextCell("C");                    // Add cell B3
            workbook.getCurrentWorksheet().addNextCell("D");                    // Add cell B4
            workbook.getCurrentWorksheet().removeCell("A2");                    // Delete cell A2
            workbook.getCurrentWorksheet().removeCell(1,1);                     // Delete cell B2
            workbook.save();                                                    // Save the workbook
            }
            catch(Exception e)
            {
                System.out.println(e.getMessage());
            } 
        }
        
        /**
         * This demo shows the usage of several styles, column widths and row heights
         */
        private static void demo4()
        {
            try
            {            
            Workbook workbook = new Workbook("test4.xlsx", "Sheet1");           // Create new workbook
            List<String> values = new ArrayList<>();                            // Create a List of values
            values.add("Header1");
            values.add("Header2");
            values.add("Header3");
            workbook.getCurrentWorksheet().addStringCellRange(values, new Cell.Address(0,0), new Cell.Address(2,0));   // Add a cell range to A4 - C4
            workbook.getCurrentWorksheet().getCells().get("A1").setStyle(BasicStyles.Bold(), workbook);                // Assign predefined basic style to cell
            workbook.getCurrentWorksheet().getCells().get("B1").setStyle(BasicStyles.Bold(), workbook);                // Assign predefined basic style to cell
            workbook.getCurrentWorksheet().getCells().get("C1").setStyle(BasicStyles.Bold(), workbook);                // Assign predefined basic style to cell
            workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 2
            workbook.getCurrentWorksheet().addNextCell(new Date(115, 9 ,1));    // Add cell A2
            workbook.getCurrentWorksheet().addNextCell(2);                      // Add cell B2
            workbook.getCurrentWorksheet().addNextCell(3);                      // Add cell B2
            workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 3
            workbook.getCurrentWorksheet().addNextCell(new Date(115, 9 ,2));    // Add cell B1
            workbook.getCurrentWorksheet().addNextCell("B");                    // Add cell B2
            workbook.getCurrentWorksheet().addNextCell("C");                    // Add cell B3

            Style s = new Style();                                              // Create new style
            s.getCurrentFill().SetColor("FF22FF11", Fill.FillType.fillColor);   // Set fill color
            s.getCurrentFont().setDoubleUnderline(true);                        // Set double underline
            s.getCurrentCellXf().setHorizontalAlign(CellXf.HorizontalAlignValue.center);  // Set alignment

            workbook.getCurrentWorksheet().getCells().get("B2").setStyle(s, workbook);    // Assign style to cell
            workbook.getCurrentWorksheet().goToNextRow();                       // Go to Row 3
            workbook.getCurrentWorksheet().addNextCell(new Date(115, 9 ,3));    // Add cell B1
            workbook.getCurrentWorksheet().addNextCell(true);                   // Add cell B2
            workbook.getCurrentWorksheet().addNextCell(false);                  // Add cell B3 
            workbook.getCurrentWorksheet().getCells().get("C2").setStyle(BasicStyles.BorderFrame(), workbook);        // Assign predefined basic style to cell

            Style s2 = new Style();                                             // Create new style
            s2.getCurrentCellXf().setTextRotation(45);                          // Set text rotation
            s2.getCurrentCellXf().setVerticalAlign(CellXf.VerticallAlignValue.center);  // Set alignment

            workbook.getCurrentWorksheet().getCells().get("B4").setStyle(s2, workbook); // Assign style to cell

            workbook.getCurrentWorksheet().setColumnWidth(0, 20f);              // Set column width
            workbook.getCurrentWorksheet().setColumnWidth(1, 15f);              // Set column width
            workbook.getCurrentWorksheet().setColumnWidth(2, 25f);              // Set column width
            workbook.getCurrentWorksheet().setRowHeight(0, 20);                 // Set row height
            workbook.getCurrentWorksheet().setRowHeight(1, 30);                 // Set row height
                      
            workbook.save();                                                    // Save the workbook
            }
            catch(Exception e)
            {
                System.out.println(e.getMessage());
            } 
        }
        
        /**
         * This demo shows the usage of cell ranges, adding and removing styles, and meta data 
         */
        private static void demo5()
        {
            try
            {   
            Workbook workbook = new Workbook("test5.xlsx", "Sheet1");           // Create new workbook
            List<String> values = new ArrayList<>();                            // Create a List of values
            values.add("Header1");
            values.add("Header2");
            values.add("Header3");
            workbook.getCurrentWorksheet().setActiveStyle(BasicStyles.BorderFrameHeader(), workbook);    // Assign predefined basic style as active style
            workbook.getCurrentWorksheet().addStringCellRange(values, "A1:C1"); // Add cell range

            values = new ArrayList<>();                                         // Create a List of values
            values.add("Cell A2");
            values.add("Cell B2");
            values.add("Cell C2");            
            workbook.getCurrentWorksheet().setActiveStyle(BasicStyles.BorderFrame(), workbook);          // Assign predefined basic style as active style
            workbook.getCurrentWorksheet().addStringCellRange(values, "A2:C2"); // Add cell range

            values = new ArrayList<>();                                         // Create a List of values
            values.add("Cell A3");
            values.add("Cell B3");
            values.add("Cell C3");            
            workbook.getCurrentWorksheet().addStringCellRange(values, "A3:C3"); // Add cell range

            values = new ArrayList<>();                                         // Create a List of values
            values.add("Cell A4");
            values.add("Cell B4");
            values.add("Cell C4");            
            workbook.getCurrentWorksheet().clearActiveStyle();                  // Clear the active style 
            workbook.getCurrentWorksheet().addStringCellRange(values, "A4:C4"); // Add cell range

            workbook.getWorkbookMetadata().setTitle("Test 5");                           // Add meta data to workbook
            workbook.getWorkbookMetadata().setSubject("This is the 5th PicoXLSX test");  // Add meta data to workbook
            workbook.getWorkbookMetadata().setCreator("PicoXLSX");                       // Add meta data to workbook
            workbook.getWorkbookMetadata().setKeywords("Keyword1;Keyword2;Keyword3");    // Add meta data to workbook

            workbook.save();                                                    // Save the workbook
            }
            catch(Exception e)
            {
                System.out.println(e.getMessage());
            } 
        }
        
        /**
         * This demo shows the usage of merging cells, protecting cells and worksheet password protection
         */
        private static void demo6()
        {
            try
            {               
            Workbook workbook = new Workbook("test6.xlsx", "Sheet1");                                        // Create new workbook
            workbook.getCurrentWorksheet().addNextCell("Mergerd1");                                          // Add cell A1
            workbook.getCurrentWorksheet().mergeCells("A1:C1");                                              // Merge cells from A1 to C1
            workbook.getCurrentWorksheet().goToNextRow();                                                    // Go to next row
            workbook.getCurrentWorksheet().addNextCell(false);                                               // Add cell A2
            workbook.getCurrentWorksheet().mergeCells("A2:D2");                                              // Merge cells from A2 to D1
            workbook.getCurrentWorksheet().goToNextRow();                                                    // Go to next row
            workbook.getCurrentWorksheet().addNextCell("22.2d");                                             // Add cell A3
            workbook.getCurrentWorksheet().mergeCells("A3:E4");                                              // Merge cells from A3 to E4
            workbook.addWorksheet("Protected");                                                              // Add a new worksheet
            workbook.getCurrentWorksheet().addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.sort);               // Allow to sort sheet (worksheet is automatically set as protected)
            workbook.getCurrentWorksheet().addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.insertRows);         // Allow to insert rows
            workbook.getCurrentWorksheet().addAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.selectLockedCells);  // Allow to select cells (locked cells caused automatically to select unlocked cells)
            workbook.getCurrentWorksheet().addNextCell("Cell A1");                                           // Add cell A1
            workbook.getCurrentWorksheet().addNextCell("Cell B1");                                           // Add cell B1
            workbook.getCurrentWorksheet().getCells().get("A1").setCellLockedState(false, true, workbook);   // Set the locking state of cell A1 (not locked but value is hidden when cell selected)
            workbook.addWorksheet("PWD-Protected");                                                          // Add a new worksheet
            workbook.getCurrentWorksheet().addCell("This worksheet is password protected. The password is:",0,0);  // Add cell A1
            workbook.getCurrentWorksheet().addCell("test123", 0, 1);                                         // Add cell A2
            workbook.getCurrentWorksheet().setSheetProtectionPassword("test123");                            // Set the password "test123"
            workbook.save();                                                                                 // Save the workbook            
            }
            catch(Exception e)
            {
                System.out.println(e.getMessage());
            }
        }        
    
}
