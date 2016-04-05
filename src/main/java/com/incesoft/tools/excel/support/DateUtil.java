package com.incesoft.tools.excel.support;

import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.regex.Pattern;

/**
 * Created by robert on 4/5/16.
 */
public class DateUtil {
    public static final int SECONDS_PER_MINUTE = 60;
    public static final int MINUTES_PER_HOUR = 60;
    public static final int HOURS_PER_DAY = 24;
    public static final int SECONDS_PER_DAY = (HOURS_PER_DAY * MINUTES_PER_HOUR * SECONDS_PER_MINUTE);

    public static final long   DAY_MILLISECONDS = SECONDS_PER_DAY * 1000L;

    /**
     * The following patterns are used in {@link #isADateFormat(int, String)}
     */
    private static final Pattern date_ptrn1 = Pattern.compile("^\\[\\$\\-.*?\\]");
    private static final Pattern date_ptrn2 = Pattern.compile("^\\[[a-zA-Z]+\\]");
    private static final Pattern date_ptrn3a = Pattern.compile("[yYmMdDhHsS]");
    private static final Pattern date_ptrn3b = Pattern.compile("^[\\[\\]yYmMdDhHsS\\-/,. :\"\\\\]+0*[ampAMP/]*$");
    //  elapsed time patterns: [h],[m] and [s]
    private static final Pattern date_ptrn4 = Pattern.compile("^\\[([hH]+|[mM]+|[sS]+)\\]");

    // variables for performance optimization:
    // avoid re-checking DataUtil.isADateFormat(int, String) if a given format
    // string represents a date format if the same string is passed multiple times.
    // see https://issues.apache.org/bugzilla/show_bug.cgi?id=55611
    private static int lastFormatIndex = -1;
    private static String lastFormatString = null;
    private static boolean cached = false;

    /**
     * Given a double, checks if it is a valid Excel date.
     *
     * @return true if valid
     * @param  value the double value
     */
    public static boolean isValidExcelDate(double value)
    {
        return (value > -Double.MIN_VALUE);
    }

    /**
     * Given a format ID this will check whether the format represents
     *  an internal excel date format or not.
     * @see #isADateFormat(int, java.lang.String)
     */
    public static boolean isInternalDateFormat(int format) {
        switch(format) {
            // Internal Date Formats as described on page 427 in
            // Microsoft Excel Dev's Kit...
            case 0x0e:
            case 0x0f:
            case 0x10:
            case 0x11:
            case 0x12:
            case 0x13:
            case 0x14:
            case 0x15:
            case 0x16:
            case 0x2d:
            case 0x2e:
            case 0x2f:
                return true;
        }
        return false;
    }

    /**
     * Given a format ID and its format String, will check to see if the
     *  format represents a date format or not.
     * Firstly, it will check to see if the format ID corresponds to an
     *  internal excel date format (eg most US date formats)
     * If not, it will check to see if the format string only contains
     *  date formatting characters (ymd-/), which covers most
     *  non US date formats.
     *
     * @param formatIndex The index of the format, eg from ExtendedFormatRecord.getFormatIndex
     * @param formatString The format string, eg from FormatRecord.getFormatString
     * @see #isInternalDateFormat(int)
     */
    public static synchronized boolean isADateFormat(int formatIndex, String formatString) {

        if (formatString != null && formatIndex == lastFormatIndex && formatString.equals(lastFormatString)) {
            return cached;
        }
        // First up, is this an internal date format?
        if(isInternalDateFormat(formatIndex)) {
            lastFormatIndex = formatIndex;
            lastFormatString = formatString;
            cached = true;
            return true;
        }

        // If we didn't get a real string, it can't be
        if(formatString == null || formatString.length() == 0) {
            lastFormatIndex = formatIndex;
            lastFormatString = formatString;
            cached = false;
            return false;
        }

        String fs = formatString;
        /*if (false) {
            // Normalize the format string. The code below is equivalent
            // to the following consecutive regexp replacements:

             // Translate \- into just -, before matching
             fs = fs.replaceAll("\\\\-","-");
             // And \, into ,
             fs = fs.replaceAll("\\\\,",",");
             // And \. into .
             fs = fs.replaceAll("\\\\\\.",".");
             // And '\ ' into ' '
             fs = fs.replaceAll("\\\\ "," ");

             // If it end in ;@, that's some crazy dd/mm vs mm/dd
             //  switching stuff, which we can ignore
             fs = fs.replaceAll(";@", "");

             // The code above was reworked as suggested in bug 48425:
             // simple loop is more efficient than consecutive regexp replacements.
        }*/
        StringBuilder sb = new StringBuilder(fs.length());
        for (int i = 0; i < fs.length(); i++) {
            char c = fs.charAt(i);
            if (i < fs.length() - 1) {
                char nc = fs.charAt(i + 1);
                if (c == '\\') {
                    switch (nc) {
                        case '-':
                        case ',':
                        case '.':
                        case ' ':
                        case '\\':
                            // skip current '\' and continue to the next char
                            continue;
                    }
                } else if (c == ';' && nc == '@') {
                    i++;
                    // skip ";@" duplets
                    continue;
                }
            }
            sb.append(c);
        }
        fs = sb.toString();

        // short-circuit if it indicates elapsed time: [h], [m] or [s]
        if(date_ptrn4.matcher(fs).matches()){
            lastFormatIndex = formatIndex;
            lastFormatString = formatString;
            cached = true;
            return true;
        }

        // If it starts with [$-...], then could be a date, but
        //  who knows what that starting bit is all about
        fs = date_ptrn1.matcher(fs).replaceAll("");
        // If it starts with something like [Black] or [Yellow],
        //  then it could be a date
        fs = date_ptrn2.matcher(fs).replaceAll("");
        // You're allowed something like dd/mm/yy;[red]dd/mm/yy
        //  which would place dates before 1900/1904 in red
        // For now, only consider the first one
        if(fs.indexOf(';') > 0 && fs.indexOf(';') < fs.length()-1) {
            fs = fs.substring(0, fs.indexOf(';'));
        }

        // Ensure it has some date letters in it
        // (Avoids false positives on the rest of pattern 3)
        if (! date_ptrn3a.matcher(fs).find()) {
            return false;
        }

        // If we get here, check it's only made up, in any case, of:
        //  y m d h s - \ / , . : [ ]
        // optionally followed by AM/PM

        boolean result = date_ptrn3b.matcher(fs).matches();
        lastFormatIndex = formatIndex;
        lastFormatString = formatString;
        cached = result;
        return result;
    }

    /**
     *  Check if a cell contains a date
     *  Since dates are stored internally in Excel as double values
     *  we infer it is a date if it is formatted as such.
     *  @see #isADateFormat(int, String)
     *  @see #isInternalDateFormat(int)
     */
    public static boolean isCellDateFormatted(double value, int styleIndex, String numFmt) {
        return DateUtil.isValidExcelDate(value)
                && isADateFormat(styleIndex, numFmt);
    }

    public static void setCalendar(Calendar calendar, int wholeDays,
                                   int millisecondsInDay, boolean use1904windowing) {
        int startYear = 1900;
        int dayAdjust = -1; // Excel thinks 2/29/1900 is a valid date, which it isn't
        if (use1904windowing) {
            startYear = 1904;
            dayAdjust = 1; // 1904 date windowing uses 1/2/1904 as the first day
        }
        else if (wholeDays < 61) {
            // Date is prior to 3/1/1900, so adjust because Excel thinks 2/29/1900 exists
            // If Excel date == 2/29/1900, will become 3/1/1900 in Java representation
            dayAdjust = 0;
        }
        calendar.set(startYear, 0, wholeDays + dayAdjust, 0, 0, 0);
        calendar.set(GregorianCalendar.MILLISECOND, millisecondsInDay);
    }

    /**
     * Get EXCEL date as Java Calendar.
     * @return Java representation of the date, or null if date is not a valid Excel date
     */
    public static Calendar getJavaCalendar(String value) {
        try {
            double date = Double.parseDouble(value);
            if (!isValidExcelDate(date)) {
                return null;
            }
            int wholeDays = (int)Math.floor(date);
            int millisecondsInDay = (int)((date - wholeDays) * DAY_MILLISECONDS + 0.5);
            Calendar calendar = new GregorianCalendar();     // using default time-zone
            setCalendar(calendar, wholeDays, millisecondsInDay, false);
            return calendar;
        } catch (Exception e) {
            return null;
        }
    }
}
