/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2017
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.style;

import java.util.List;

/**
 * Class representing a collection of Sub-Styles 
 * @author Raphael Stoeckli
 */ 
public class StyleCollection
   {
       private List<Border> borders;
       private List<Fill> fills;
       private List<Font> fonts;
       private List<NumberFormat> numberFormats;
       private List<CellXf> cellXfs;

       /**
        * Gets the list of border
        * @return List of border
        */
        public List<Border> getBorders() {
            return borders;
        }

        /**
         * Sets the list of border
         * @param borders List of border 
         */
        public void setBorders(List<Border> borders) {
            this.borders = borders;
        }

        /**
         * Gets the list of fills
         * @return List of fills
         */
        public List<Fill> getFills() {
            return fills;
        }

        /**
         * Sets the list of fills
         * @param fills List of fills
         */
        public void setFills(List<Fill> fills) {
            this.fills = fills;
        }

        /**
         * Gets the list of fonts
         * @return List of fonts
         */
        public List<Font> getFonts() {
            return fonts;
        }

        /**
         * Sets the list of fonts
         * @param fonts List of fonts
         */
        public void setFonts(List<Font> fonts) {
            this.fonts = fonts;
        }

        /**
         * Gets the list of number formats
         * @return List of number formats
         */
        public List<NumberFormat> getNumberFormats() {
            return numberFormats;
        }

        /**
         * Sets the list of number formats
         * @param numberFormats List of number formats
         */
        public void setNumberFormats(List<NumberFormat> numberFormats) {
            this.numberFormats = numberFormats;
        }

        /**
         * Gets the list of CellXFs
         * @return List of CellXFs
         */
        public List<CellXf> getCellXfs() {
            return cellXfs;
        }

        /**
         * Sets the list of CellXFs
         * @param cellXfs List of CellXFs
         */
        public void setCellXfs(List<CellXf> cellXfs) {
            this.cellXfs = cellXfs;
        }
       
        /**
         * Default constructor
         */
        public StyleCollection()
        {
            
        }
       
   }
