/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2015
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j;

import java.util.Calendar;
import java.util.Date;

/**
 * Class for shared used (static) methods
 * @author Raphael Stoeckli
 */
public class Helper {
    
    /**
     * Method of a string to check whether its reference is null or the content is empty
     * @param value value / reference to check
     * @return True if the passed parameter is null or empty, otherwise false
     */
    public static boolean isNullOrEmpty(String value)
    {
        if (value == null) { return true; }
        if (value.isEmpty() == true ){ return true; }
        else { return false; }
    } 
    
    /**
     * Method to calculate the OA date (OLE automation) of the passed date.<br>
     * The date is the number of days since the 01.01.1900 00:00. The hours of a day is a float between 0 and 1
     * @param date Date to convert
     * @return OA date
     */
    public static double getOADate(Date date)
    {
        Calendar root = Calendar.getInstance();
        Calendar dateCal = Calendar.getInstance();
        root.set(1899, 11, 29);
        dateCal.setTime(date);
        long t1 = root.getTimeInMillis();
        long t2 = dateCal.getTimeInMillis();
        double span = t2 - t1;
        double days = Math.floor(span / (1000 * 24 * 60 * 60));
        double h = dateCal.get(Calendar.HOUR_OF_DAY);
        double m = dateCal.get(Calendar.MINUTE);
        double s = dateCal.get(Calendar.SECOND);
        return ((s + (m * 60) + (h * 3600)) / 86400) + days;
    }    
   
    
}
