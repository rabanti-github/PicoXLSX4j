/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2017
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package picoxlsx4j;

import java.util.Calendar;
import java.util.Date;
import picoxlsx4j.exception.FormatException;

/**
 * Class for shared used (static) methods
 * @author Raphael Stoeckli
 */
public class Helper {
        
// ### C O N S T A N T S ###    
    //private static Calendar root = Calendar.getInstance();
    //static { root.set(1899, 11, 29); }
    private static long rootTicks;
    static
    {
        Calendar rootCalendar = Calendar.getInstance();
        rootCalendar.set(1899, 11, 29);
        rootTicks = rootCalendar.getTimeInMillis();
    }
    
// ### S T A T I C   M E T H O D S ###    
    /**
     * Method to calculate the OA date (OLE automation) of the passed date.<br>
     * OA Date format starts at January 1st 1900 (actually 00.01.1900). Dates beyond this date cannot be handled by Excel under normal circumstances and will throw a FormatException
     * @param date Date to convert
     * @exception FormatException Throws a FormatException if the passed date cannot be translated to the OADate format
     * @return OA date
     */
    public static double getOADate(Date date)
    {
       // Calendar root = Calendar.getInstance();
        Calendar dateCal = Calendar.getInstance();
        dateCal.setTime(date);
        //long t1 = root.getTimeInMillis();
        //long t2 = dateCal.getTimeInMillis();
        long currentTicks = dateCal.getTimeInMillis();
        /*
        double span = t2 - t1;
        double days = Math.floor(span / (86400000)); // 1000 * 24 * 60 * 60
        double h = dateCal.get(Calendar.HOUR_OF_DAY);
        double m = dateCal.get(Calendar.MINUTE);
        double s = dateCal.get(Calendar.SECOND);
        return ((s + (m * 60) + (h * 3600)) / 86400) + days;
        */
        double d = ((dateCal.get(Calendar.SECOND) + (dateCal.get(Calendar.MINUTE) * 60) + (dateCal.get(Calendar.HOUR_OF_DAY) * 3600)) / 86400) + Math.floor((currentTicks - rootTicks) / (86400000));
        if (d < 0)
        {
            throw new FormatException("The date is not in a valid range for Excel. Dates before 1900-01-01 are not allowed.");
        }
        return d;
    }    
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
   
    
}
