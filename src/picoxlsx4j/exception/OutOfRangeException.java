/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2016
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j.exception;

/**
 * Class for exceptions regarding out-of-range incidents
 * @author Raphael Stoeckli
 */
public class OutOfRangeException extends RuntimeException{
 
    /**
     * Default constructor
     */    
    public OutOfRangeException()
    {
        super();
    }
    
    /**
     * Constructor with passed message
     * @param message Message of the exception
     */    
    public OutOfRangeException(String message)
    {
        super(message);
    }
    
    
}
