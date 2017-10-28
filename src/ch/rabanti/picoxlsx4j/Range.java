/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2017
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.picoxlsx4j;

/**
 * Class representing a cell range (used like a simple struct)
 * @author Raphael Stoeckli
 */
    public class Range
    {
// ### P U B L I C  F I E L D S ###        
        /**
         * End address of the range
         */
        public final Address EndAddress;
        /**
         * Start address of the range
         */
        public final Address StartAddress;
        
// ### C O N S T R U C T O R S ###        
        /**
         * Constructor with parameters
         * @param start Start address of the range
         * @param end End address of the range
         */
        public Range(Address start, Address end)
        {
            this.StartAddress = start;
            this.EndAddress = end;
        }
        
// ### M E T H O D S ###        
        /**
         * Overwritten toString method
         * @return Returns the range (e.g. 'A1:B12')
         */
        @Override
        public String toString()
        {
            return StartAddress.toString() + ":" + EndAddress.toString();
        }
        
    } 