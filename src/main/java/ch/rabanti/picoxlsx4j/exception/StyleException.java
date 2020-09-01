/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2020
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.picoxlsx4j.exception;

/**
 * Class for exceptions regarding Styles
 * @author Raphael Stoeckli
 */
public class StyleException extends RuntimeException{
    
    private String exceptionTitle;
    
    /**
     * Gets the title of the exception
     * @return Title as string
     */
    public String getExceptionTitle() {
        return this.exceptionTitle;
    }
    
    
    /**
     * Default constructor
     */    
    public StyleException()
    {
        super();
    }
    
    /**
     * Constructor with passed message
     * @param title Title of the exception
     * @param message Message of the exception
     */    
    public StyleException(String title, String message)
    {
        super(title + ": " + message);
        this.exceptionTitle = title;
    }
    
    
}
