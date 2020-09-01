/*
 * PicoXLSX4j is a small Java library to generate XLSX (Microsoft Excel 2007 or newer) files in an easy and native way
 * Copyright Raphael Stoeckli © 2020
 * This library is licensed under the MIT License.
 * You find a copy of the license in project folder or on: http://opensource.org/licenses/MIT
 */
package ch.rabanti.picoxlsx4j;

import ch.rabanti.picoxlsx4j.exception.FormatException;
import ch.rabanti.picoxlsx4j.lowLevel.LowLevel;

import java.time.LocalTime;
import java.util.Calendar;
import java.util.Date;

/**
 * Class for shared used (static) methods
 * @author Raphael Stoeckli
 */
public class Helper {
        
// ### C O N S T A N T S ###    
    private static final long ROOT_TICKS;

    /**
     * Minimum valid OAdate value (1900-01-01)
     */
    public static final double MIN_OADATE_VALUE = 0f;
    /**
     * Maximum valid OAdate value (9999-12-31)
     */
    public static final double MAX_OADATE_VALUE = 2958465.9999f;

    static
    {
        Calendar rootCalendar = Calendar.getInstance();
        rootCalendar.set(1899, Calendar.DECEMBER, 30,0,0,0);
        ROOT_TICKS = rootCalendar.getTimeInMillis();
    }
    
// ### S T A T I C   M E T H O D S ###    
    /**
     * Method to calculate the OA date (OLE automation) of the passed date.<br>
     * OA Date format starts at January 1st 1900 (actually 00.01.1900)and ends at December 31 9999. Values beyond these dates cannot be handled by Excel under normal circumstances and will throw a FormatException
     * @param date Date to convert
     * @exception FormatException Throws a FormatException if the passed date cannot be translated to the OADate format
     * @return Date or date and time as number
     */
    public static String getOADateTimeString(Date date)
    {
        Calendar dateCal = Calendar.getInstance();
        dateCal.setTime(date);
        double currentTicks = dateCal.getTimeInMillis();
        double d = ((double)(dateCal.get(Calendar.SECOND) + (dateCal.get(Calendar.MINUTE) * 60) + (dateCal.get(Calendar.HOUR_OF_DAY) * 3600)) / 86400) + Math.floor((currentTicks - ROOT_TICKS) / (86400000));
        if (d < MIN_OADATE_VALUE || d > MAX_OADATE_VALUE)
        {
            throw new FormatException("FormatException","The date is not in a valid range for Excel. Dates before 1900-01-01 are not allowed.");
        }
        return  Double.toString(d);
    }

    /**
     * Method to convert a time into the internal Excel time format (OAdate without days)
     * <p>The time is represented by a OAdate without the date component. A time range is between &gt;0.0 (00:00:00) and &lt;1.0 (23:59:59)</p>
     * @param time Time to process.
     * @exception FormatException Throws a FormatException if the passed time cannot be translated to the OADate format
     * @return Time as number
     */
    public static String getOATimeString(LocalTime time)
    {
        try {
            int seconds = time.getSecond() + time.getMinute() * 60 + time.getHour() * 3600;
            double d = (double)seconds / 86400d;
            return Double.toString(d);
        }
        catch (Exception ex){
            throw new FormatException("ConversionException","The time could not be transformed into Excel format (OADate).", ex);
        }
    }

    /**
     * Method of a string to check whether its reference is null or the content is empty
     * @param value value / reference to check
     * @return True if the passed parameter is null or empty, otherwise false
     */
    public static boolean isNullOrEmpty(String value)
    {
        if (value == null) { return true; }
        return value.isEmpty();
    }
   
    
}
