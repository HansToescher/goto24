package at.goto24.data.util;


import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.TimeZone;
import java.util.concurrent.ConcurrentHashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.namespace.QName;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;



public final class DataFormatter {
	private static final Logger LOG = Logger.getLogger(DataFormatter.class);

    public static final String DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";

    public static final String MONTH_DAY_FORMAT = "dd.MM";

    public static final String FILENAME_DATE_FORMAT = "yyyy-MM-dd_HH-mm-ss";

    public static final String SOCCERWAY_DATEFORMAT = "yyyy-MM-dd";

    public static final String ORACLE_DATEFORMAT = "yyyy-MM-dd HH:mm:ss";

    public static final String DATE_FORMAT_DAY = "yyyy_MM_dd";

    public static final String DATE_FORMAT_MONTHLY = "yyyy_MM";

    public static final String DATE_FORMAT_MONTH = "yyyy-MM";

    public static final String DATE_TIME_FORMAT = "yyyy-MM-dd HH:mm:ss.SSS";

    public static final String DATE_FORMAT_VERY_LONG = "EEE MMM dd HH:mm:ss.SSS z yyyy";

    public static final String STRANGE_FORMAT = "MM/dd/yyyy";
    
    public static final String REGULAR_DATE_FORMAT= "dd/MM/yyyy";

    public static final String BIRTHDAY_FORMAT = "dd.MM.yyyy";

    public static final String BIRTHDAY_FORMAT_MONTH = "MM.yyyy";

    public static final String YEAR_FORMAT = "yyyy";

    public static final String YEAR_MONTH_FORMAT = "MM/yyyy";

    public static final String TIME_FORMAT = "HH:mm:ss";

    public static final String DATE_FORMAT_MINUTE = "dd.MM.yyyy HH:mm";

    public static final String DATE_FORMAT_WEEKDAY = "EEE";

    public static final String VISUALIZATION_DATE_FORMAT = "yyyy/MM/dd";

    public static final long HOUR = 1000 * 60 * 60;

    public static final long DAY = HOUR * 24;

    public static final long MONTH = 30 * DAY;

    private final Map<Integer, Map<String, Map<String, SimpleDateFormat>>> dateFormats = new ConcurrentHashMap<Integer, Map<String, Map<String, SimpleDateFormat>>>();

    public static final String XML_HEADER = "<?xml version='1.0' encoding='UTF-8'?>";

    private static DataFormatter instance = null;

    private static final Map<String, String> DAYS = new HashMap<String, String>();

    private static final Object OBJ = new Object();

	//Regex for checking if string ends with a DOt and 3 digits (for date set by OWS to set miliseconds).
	private final static String REGEX_ENDS_WITH_MILLISECONDS = "\\.\\d{3}$";

    private DataFormatter() {
	DAYS.put("Mo", "Mon");
	DAYS.put("Di", "Tue");
	DAYS.put("Mi", "Wed");
	DAYS.put("Do", "Thu");
	DAYS.put("Fr", "Fri");
	DAYS.put("Sa", "Sat");
	DAYS.put("So", "Sun");
    }

    public static DataFormatter getInstance() {
	synchronized (OBJ) {
	    if (instance == null) {
		instance = new DataFormatter();
	    }
	}
	return instance;
    }

    public String getMonthStartDateAsString(Date date) {
	Calendar c = Calendar.getInstance(TimeZone.getTimeZone("GMT"));
	c.setTime(date);
	c.set(Calendar.DATE, 1);
	c.set(Calendar.HOUR_OF_DAY, 0);
	c.set(Calendar.MINUTE, 0);
	c.set(Calendar.SECOND, 0);
	c.set(Calendar.MILLISECOND, 0);

	return getDateFormat(DATE_FORMAT).format(c.getTime());
    }

    public Date getMonthStartDate(Date date) {
	Calendar c = Calendar.getInstance(TimeZone.getTimeZone("GMT"));
	c.setTime(date);
	c.set(Calendar.DATE, 1);
	c.set(Calendar.HOUR_OF_DAY, 0);
	c.set(Calendar.MINUTE, 0);
	c.set(Calendar.SECOND, 0);
	c.set(Calendar.MILLISECOND, 0);

	return c.getTime();
    }

    public SimpleDateFormat getDateFormat(String dateFormat) {
	return getDateFormat(dateFormat, "GMT");
    }

    public SimpleDateFormat getDateFormat(String dateFormat, String timeZone) {
	Map<String, Map<String, SimpleDateFormat>> dfMap = dateFormats.get(Thread.currentThread().hashCode());
	if (dfMap == null) {
	    dfMap = new ConcurrentHashMap<String, Map<String, SimpleDateFormat>>();
	    dateFormats.put(Thread.currentThread().hashCode(), dfMap);
	}

	Map<String, SimpleDateFormat> dfTzMap = dfMap.get(timeZone);
	if (dfTzMap == null) {
	    dfTzMap = new ConcurrentHashMap<String, SimpleDateFormat>();
	    dfMap.put(timeZone, dfTzMap);
	}

	SimpleDateFormat df = dfTzMap.get(dateFormat);
	if (df == null) {
	    df = new SimpleDateFormat(dateFormat);
	    df.setTimeZone(TimeZone.getTimeZone(timeZone));
	    dfTzMap.put(dateFormat, df);
	}

	return df;
    }

    public String removeSpaces(String input) {
	boolean lastIsSpace = false;
	char[] array = input.toCharArray();
	StringBuilder sb = new StringBuilder();
	for (int i = 0; i < array.length; i++) {
	    if (array[i] == ' ') {
		if (!lastIsSpace) {
		    sb.append(array[i]);
		    lastIsSpace = true;
		}
	    } else {
		sb.append(array[i]);
		lastIsSpace = false;
	    }
	}
	return sb.toString();
    }

    public String removeAllSpaces(String input) {
	String[] array = input.split(" ");
	StringBuilder sb = new StringBuilder();
	for (int i = 0; i < array.length; i++) {
	    sb.append(array[i]);
	}
	return sb.toString();
    }

    public boolean isNumber(String s) {
	String validChars = "0123456789-";
	boolean isNumber = true;

	if (s == null || s.length() == 0 || "".equals(s)) {
	    return false;
	}

	for (int i = 0; i < s.length() && isNumber; i++) {
	    char c = s.charAt(i);
	    if (validChars.indexOf(c) == -1) {
		isNumber = false;
	    } else {
		isNumber = true;
	    }
	}
	return isNumber;
    }

    public boolean isDoubleNumber(String s) {
	String validChars = "0123456789.";
	boolean isNumber = true;

	if (s == null || s.length() == 0) {
	    return false;
	}

	for (int i = 0; i < s.length() && isNumber; i++) {
	    char c = s.charAt(i);
	    if (validChars.indexOf(c) == -1) {
		isNumber = false;
	    } else {
		isNumber = true;
	    }
	}
	return isNumber;
    }

    public static void deleteXMLNamespace(XmlObject x) {
	String s;
	XmlCursor c = x.newCursor();
	c.toNextToken();
	while (c.hasNextToken()) {
	    if (c.isNamespace()) {
		c.removeXml();
	    } else {
		if (c.isStart() || c.isAttr()) {
		    s = c.getName().getLocalPart();
		    c.setName(new QName(s));
		}
		c.toNextToken();
	    }
	}
	c.dispose();
    }

    public static String getLongerWeekday(String day) {
	String result = day;
	for (String key : DAYS.keySet()) {
	    result = result.replace(key, DAYS.get(key));
	}

	return result;
    }

    // returns start and end date of the day before the given refDate in the
    // correct timezone from 00:00:00 to 23:59:59
    public static Date[] getEntireDayBefore(Date refDate, String timezoneName) {
	Date[] result = new Date[2];

	Calendar cal = Calendar.getInstance(TimeZone.getTimeZone(timezoneName));
	cal.setTime(refDate);
	cal.add(Calendar.DAY_OF_YEAR, -1);
	cal.set(Calendar.HOUR_OF_DAY, 0);
	cal.set(Calendar.MINUTE, 0);
	cal.set(Calendar.SECOND, 0);
	cal.set(Calendar.MILLISECOND, 0);
	result[0] = cal.getTime();
	cal.add(Calendar.DAY_OF_YEAR, 1);
	cal.add(Calendar.SECOND, -1);
	result[1] = cal.getTime();

	return result;
    }

    // returns start and end date of the week before the given refDate in the
    // correct timezone from Monday 00:00:00 to Sunday 23:59:59
    public static Date[] getEntireWeekBefore(Date refDate, String timezoneName) {
	Date[] result = new Date[2];

	Calendar cal = Calendar.getInstance(TimeZone.getTimeZone(timezoneName));
	cal.setTime(refDate);
	cal.add(Calendar.WEEK_OF_YEAR, -1);
	cal.set(Calendar.DAY_OF_WEEK, Calendar.MONDAY);
	cal.set(Calendar.HOUR_OF_DAY, 0);
	cal.set(Calendar.MINUTE, 0);
	cal.set(Calendar.SECOND, 0);
	cal.set(Calendar.MILLISECOND, 0);
	result[0] = cal.getTime();
	cal.add(Calendar.WEEK_OF_YEAR, 1);
	cal.add(Calendar.SECOND, -1);
	result[1] = cal.getTime();

	return result;
    }

    // returns start and end date of the month before the given refDate in the
    // correct timezone from 1st of month 00:00:00 to last of month 23:59:59
    public static Date[] getEntireMonthBefore(Date refDate, String timezoneName) {
	Date[] result = new Date[2];

	Calendar cal = Calendar.getInstance(TimeZone.getTimeZone(timezoneName));
	cal.setTime(refDate);
	cal.add(Calendar.MONTH, -1);
	cal.set(Calendar.DAY_OF_MONTH, 1);
	cal.set(Calendar.HOUR_OF_DAY, 0);
	cal.set(Calendar.MINUTE, 0);
	cal.set(Calendar.SECOND, 0);
	cal.set(Calendar.MILLISECOND, 0);
	result[0] = cal.getTime();
	cal.add(Calendar.MONDAY, 1);
	cal.add(Calendar.SECOND, -1);
	result[1] = cal.getTime();

	return result;
    }

    public static Date getMondayOfLastWeek(Date reportDate) {
	int dayDiff = 0;
	Calendar cal = Calendar.getInstance();
	cal.setTime(reportDate);
	int currentWeekDay = cal.get(Calendar.DAY_OF_WEEK);
	if (currentWeekDay == Calendar.SUNDAY) {
	    dayDiff = 13;
	} else {
	    dayDiff = currentWeekDay + 5;
	}

	return new Date(reportDate.getTime() - (DataFormatter.DAY * dayDiff));
    }

    public static String escapeForFilename(String input) {
	if (input != null) {
	    String result = input;
	    result = result.replaceAll("Ä", "Ae");
	    result = result.replaceAll("ä", "ae");
	    result = result.replaceAll("Ö", "Oe");
	    result = result.replaceAll("ö", "oe");
	    result = result.replaceAll("Ü", "Ue");
	    result = result.replaceAll("ü", "ue");
	    result = result.replaceAll("ß", "ss");
	    result = result.replaceAll("\\.", "");
	    result = result.replaceAll("\\,", "");
	    result = result.replaceAll("/", " ");

	    return result;
	}

	return "";
    }

    public double getConvertedTime(Date date, String timeZone, String dateFormat) {
	SimpleDateFormat timeFormat = DataFormatter.getInstance().getDateFormat(dateFormat, timeZone);
	return HSSFDateUtil.convertTime(timeFormat.format(date));
    }

    public double getConvertedDate(Date date, String timeZone) {
	Calendar cal = Calendar.getInstance(TimeZone.getTimeZone(timeZone));
	cal.setTime(date);
	Calendar tmp = Calendar.getInstance(TimeZone.getTimeZone(timeZone));
	double totalDays = 1d;
	for (int i = 1900; i < cal.get(Calendar.YEAR); i++) {
	    tmp.set(Calendar.YEAR, i);
	    int daysForThatYear = tmp.getActualMaximum(Calendar.DAY_OF_YEAR);
	    totalDays += daysForThatYear;
	}
	totalDays += cal.get(Calendar.DAY_OF_YEAR);

	return totalDays;
    }

	/**
	 * It checks if string ends with [DOT, NUMBER, NUMBER, NUMBER] in case it does not suffix is added.
	 *
	 * @param dateString test to be checked
	 * @param suffix     to be added
	 * @return Empty String in case of null date string input OR  input string + suffix.
	 */
	public static String checkForAndSetMilliseconds(String dateString, String suffix) {
		StringBuilder result = new StringBuilder();

		if (dateString != null) {
			result.append(dateString);

			Pattern pattern = Pattern.compile(REGEX_ENDS_WITH_MILLISECONDS);
			Matcher matcher = pattern.matcher(dateString);

			if (!matcher.find()) {
				result.append(suffix);
				if (LOG.isDebugEnabled()) {
					LOG.debug("Date was corrected from [" + dateString + "] to [" + result + "]");
				}
			}
		}

		return result.toString();
	}
}
