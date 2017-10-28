/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2017
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.picoxlsx4j;

import static ch.rabanti.picoxlsx4j.Cell.resolveCellAddress;
import ch.rabanti.picoxlsx4j.exception.RangeException;

/**
 * C class representing a cell address (column and row; used like a simple struct)
 * @author Raphael Stoeckli
 */
    public class Address
    {
       
// ### P U B L I C  F I E L D S ###        
        /**
         * Column of the address (zero-based)
         */        
        public final int Column;
        /**
         * Row of the address (zero-based)
         */
        public final int Row;
        
// ### C O N S T R U C T O R S ###        
        /**
         * Constructor with parameters
         * @param column Column of the address (zero-based)
         * @param row Row of the address (zero-based)
         */
        public Address(int column, int row)
        {
            this.Column = column;
            this.Row = row;
        }

// ### M E T H O D S ###
        /**
         * Gets the address as string
         * @return Address as string
         * @throws RangeException Thrown if the column or row is out of range
         */
        public String getAddress()
        {
            return resolveCellAddress(this.Column, this.Row);
        }
        
        /**
         * Returns the address as string or "Illegal Address" in case of an exception
         * @return Address or notification in case of an error
         */
        @Override
        public String toString()
        {
            try
            {
            return getAddress();
            }
            catch(Exception e)
            {
                return "Illegal Address";
            }
        }
    }