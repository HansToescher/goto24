package at.goto24.data.util;

import com.mchange.v2.c3p0.ComboPooledDataSource;

import at.goto24.data.ItemRecord;
import at.goto24.data.generic.Cf;

import java.io.File;
import java.sql.*;
import java.text.DateFormatSymbols;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Date;
import java.util.Map.Entry;

import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.log4j.Logger;
import org.w3c.dom.Element;



public final class Util {
  private static final Logger LOG = Logger.getLogger(Util.class);
    
  public final static String REGEX_DECIMAL = "^[-+]?\\d{1,10}(?:[\\.][1-9]\\d{0,9})?$";    	    
    
  public final static String REGEX_BOOLEAN = "^(true|false){1,1}$";

  public static final String ITEMTYPE = "ITEMTYPE";
  public static final String NAME = "NAME";
  

  enum DateFunction { ADD, SUB }
    
    private Util() {
    }

    private static boolean checkRegEx(String pattern, String text) {
        if (text != null) {
            return text.toLowerCase().matches(pattern);
        }
        return false;
    }

    /** Case insensitive check if text is a decimal number.
     * 
     * As decimal separator  dot .] is valid.
     * 10 digits before or/and after decimal separators are OK. 
     * 
     * @param text
     * @return
     */
    public static boolean isDecimalNumber(String text) {
        return checkRegEx(Util.REGEX_DECIMAL, text);
    }
    
	/**
	 * Case insensitive check if text is 'true' or 'false' literally.
	 * 
	 * @param text
	 * @return true in case of "true" or "false" text else false.
	 */
	public static boolean isBoolean(String text) {
		return checkRegEx(Util.REGEX_BOOLEAN, text);
	}

    public static String convertFileSeparators(String inputFilename) {
        String result = inputFilename;
        try {
            while (result.contains("\\\\")) {
                result = result.replace("\\\\", File.separator);
            }
            while (result.contains("//")) {
                result = result.replace("//", File.separator);
            }
            if (File.separator.equals("\\")) {
                while (result.contains("/")) {
                    result = result.replace("/", File.separator);
                }
            }
            if (File.separator.equals("/")) {
                while (result.contains("\\")) {
                    result = result.replaceAll("\\", File.separator);
                }
            }
        } catch (Exception e) {
            LOG.error("error", e);
        }
        return result;
    }

    public static void debugXML(Element elem) {
        try {
            TransformerFactory tFactory = TransformerFactory.newInstance();
            Transformer transformer = tFactory.newTransformer();
            DOMSource source = new DOMSource(elem);
            StreamResult result = new StreamResult(System.out);
            transformer.transform(source, result);
        } catch (Exception e) {
            LOG.error("error", e);
        }
    }

    /*
     * depending on the date passed with reportDate and the timezone a String
     * date is returned. example:
     * date="REPORTDATE_DB, reportdate="2011-02-21", timeZoneId="
     * Asia/Singapore(GMT+8)" --> result="2011-02-20 16:00:00"
     */
    public static Timestamp convertString(String date, Date reportdate, String timeZoneId) throws ParseException{

        TimeZone tzw = TimeZone.getTimeZone(timeZoneId);
        
        Timestamp result = null;
        
        if (date!=null)
        {
	        if (valueIsDateTime(date)) {
	        	String[] aElems = date.split("[.]");
	        	SimpleDateFormat sdf;
	        	if (aElems.length == 1) {
	        		sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	        	} else {
	        		sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
	        	}
	        	Date parDate = sdf.parse(date);
	        	Timestamp tsParDate = new Timestamp(parDate.getTime());
	        	return tsParDate;
	        }
	        
	        
	        if (date.indexOf(":") > -1) { // defined Date detected
	        	SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	        	SimpleDateFormat sdf16 = new SimpleDateFormat("yyyy-MM-dd HH:mm");
	        	Date parDate = null;
	        	if (date.length() == 16) {
	        		parDate = sdf16.parse(date);
	        	} else {
	        		parDate = sdf.parse(date);
	        	}
	        	
	        	Timestamp tsParDate = new Timestamp(parDate.getTime());
	        	return tsParDate;
	        } 
	        
	        StringTokenizer tok = new StringTokenizer(date, "-+", true);
	
	        Calendar dvalue = Calendar.getInstance();
	        dvalue.setTimeZone(tzw);
	
	        while (tok.hasMoreElements()) {
	            String val = tok.nextToken();
	
	            if ("REPORTDATE".equalsIgnoreCase(val)) {
	                dvalue.setTime(reportdate);
	            } else if ("REPORTDATE_DB".equalsIgnoreCase(val)) {
	                dvalue.setTime(reportdate);
	
	                dvalue.set(Calendar.HOUR_OF_DAY, 0);
	                dvalue.set(Calendar.MINUTE, 0);
	                dvalue.set(Calendar.SECOND, 0);
	                dvalue.set(Calendar.MILLISECOND, 0);
	                
	            } else if ("REPORTDATE_DE".equalsIgnoreCase(val)) {
	                dvalue.setTime(reportdate);
	                dvalue.set(Calendar.HOUR_OF_DAY, 23);
	                dvalue.set(Calendar.MINUTE, 59);
	                dvalue.set(Calendar.SECOND, 59);
	                dvalue.set(Calendar.MILLISECOND, 997);
	
	            } else if ("REPORTDATE_MB".equalsIgnoreCase(val)) {
	                dvalue.setTime(reportdate);
	                dvalue.set(Calendar.HOUR_OF_DAY, 0);
	                dvalue.set(Calendar.MINUTE, 0);
	                dvalue.set(Calendar.SECOND, 0);
	                dvalue.set(Calendar.MILLISECOND, 0);
	                dvalue.set(Calendar.DAY_OF_MONTH, 1);  // First day of current month
	
	            } else if ("REPORTDATE_ME".equalsIgnoreCase(val)) {
	                dvalue.setTime(reportdate);
	                dvalue.set(Calendar.HOUR_OF_DAY, 23);
	                dvalue.set(Calendar.MINUTE, 59);
	                dvalue.set(Calendar.SECOND, 59);
	                dvalue.set(Calendar.MILLISECOND, 997);
	                dvalue.set(Calendar.DAY_OF_MONTH, 1);  // First day of current month
	                dvalue.add(Calendar.MONTH, 1);         // First day of next month
	                dvalue.add(Calendar.DAY_OF_MONTH, -1); // Previous day -> Last day of current month
	
	            } else if ("REPORTDATE_LMB".equalsIgnoreCase(val)) {
	                dvalue.setTime(reportdate);
	                dvalue.set(Calendar.HOUR_OF_DAY, 0);
	                dvalue.set(Calendar.MINUTE, 0);
	                dvalue.set(Calendar.SECOND, 0);
	                dvalue.set(Calendar.MILLISECOND, 0);
	                dvalue.set(Calendar.DAY_OF_MONTH, 1);  // First day of current month
	                dvalue.add(Calendar.MONTH, -1);        // First day of previous month
	
	            } else if ("REPORTDATE_LME".equalsIgnoreCase(val)) {
	                dvalue.setTime(reportdate);
	                dvalue.set(Calendar.HOUR_OF_DAY, 23);
	                dvalue.set(Calendar.MINUTE, 59);
	                dvalue.set(Calendar.SECOND, 59);
	                dvalue.set(Calendar.MILLISECOND, 997);
	                dvalue.set(Calendar.DAY_OF_MONTH, 1);  // First day of current month
	                dvalue.add(Calendar.DAY_OF_MONTH, -1); // Previous day -> Last day of previous month
	
	            } else if ("REPORTDATE_WB".equalsIgnoreCase(val)) {
	                dvalue.setTime(reportdate);
	                dvalue.set(Calendar.DAY_OF_WEEK, 2); // Monday
	                if (dvalue.getTimeInMillis() > reportdate.getTime()) {
	                    dvalue.add(Calendar.WEEK_OF_MONTH, -1);
	                }
	                dvalue.set(Calendar.HOUR_OF_DAY, 0);
	                dvalue.set(Calendar.MINUTE, 0);
	                dvalue.set(Calendar.SECOND, 0);
	                dvalue.set(Calendar.MILLISECOND, 0);
	            } else if ("REPORTDATE_WE".equalsIgnoreCase(val)) {
	                dvalue.setTime(reportdate);
	                dvalue.set(Calendar.DAY_OF_WEEK, 1); // Sunday
	                if (dvalue.getTimeInMillis() < reportdate.getTime()) {
	                    dvalue.add(Calendar.WEEK_OF_MONTH, 1);
	                }
	                dvalue.set(Calendar.HOUR_OF_DAY, 23);
	                dvalue.set(Calendar.MINUTE, 59);
	                dvalue.set(Calendar.SECOND, 59);
	                dvalue.set(Calendar.MILLISECOND, 997);    
	            } else if ("REPORTDATE_YB".equalsIgnoreCase(val)) {
	            	dvalue.setTime(reportdate);
	            	dvalue.set(Calendar.MONTH, Calendar.JANUARY);
	            	dvalue.set(Calendar.HOUR_OF_DAY, 0);
	            	dvalue.set(Calendar.MINUTE, 0);
	                dvalue.set(Calendar.SECOND, 0);
	                dvalue.set(Calendar.MILLISECOND, 0);
	                dvalue.set(Calendar.DAY_OF_MONTH, 1);
	            } else if ("REPORTDATE_YE".equalsIgnoreCase(val)) {
	            	dvalue.setTime(reportdate);
	            	dvalue.set(Calendar.MONTH, Calendar.DECEMBER);
	            	dvalue.set(Calendar.HOUR_OF_DAY, 23);
	            	dvalue.set(Calendar.MINUTE, 59);
	                dvalue.set(Calendar.SECOND, 59);
	                dvalue.set(Calendar.MILLISECOND, 997);
	                dvalue.set(Calendar.DAY_OF_MONTH, 31);
	            } else if ("REPORTDATE_LYB".equalsIgnoreCase(val)) {
	            	dvalue.setTime(reportdate);
	                int year =  dvalue.get(Calendar.YEAR);
	        	    dvalue.set(Calendar.YEAR, year - 1);
	                dvalue.set(Calendar.DAY_OF_YEAR, 1);
	                dvalue.set(Calendar.HOUR_OF_DAY, 0);
	                dvalue.set(Calendar.MINUTE, 0);
	                dvalue.set(Calendar.SECOND, 0);
	                dvalue.set(Calendar.MILLISECOND, 0);
	            } else if ("REPORTDATE_LYE".equalsIgnoreCase(val)) {
	            	dvalue.setTime(reportdate);
	            	int year =  dvalue.get(Calendar.YEAR);
	            	dvalue.set(Calendar.YEAR, year - 1);
	            	dvalue.set(Calendar.MONTH, Calendar.DECEMBER);
	            	dvalue.set(Calendar.HOUR_OF_DAY, 23);
	            	dvalue.set(Calendar.MINUTE, 59);
	                dvalue.set(Calendar.SECOND, 59);
	                dvalue.set(Calendar.MILLISECOND, 997);
	                dvalue.set(Calendar.DAY_OF_MONTH, 31);
	            } else if ("REPORTDATE_HS".equalsIgnoreCase(val)) {
	            	dvalue.setTime(reportdate);
	            	dvalue.set(Calendar.MINUTE, 0);
	                dvalue.set(Calendar.SECOND, 0);
	                dvalue.set(Calendar.MILLISECOND, 0);
	            }
	            else if ("REPORTDATE_HE".equalsIgnoreCase(val))
	            {
	            	dvalue.setTime(reportdate);
	            	dvalue.set(Calendar.MINUTE, 59);
	                dvalue.set(Calendar.SECOND, 59);
	                dvalue.set(Calendar.MILLISECOND, 997);
	            }
	            else if ("+".equalsIgnoreCase(val)) {
	                changeDateTimePart(dvalue, tok.nextToken(), DateFunction.ADD);
	
	            } else if ("-".equalsIgnoreCase(val)) {
	                changeDateTimePart(dvalue, tok.nextToken(), DateFunction.SUB);
	
	            }
	            result = new Timestamp(dvalue.getTimeInMillis());
	        }
        }

        LOG.debug("Converted " + date + " to " + (result!=null ? result.toString() : "null"));

        return result;
    }


    public static String convert2SqlInString(Collection<String> input) {
        Collection<String> result = new ArrayList<String>();
        for (String s : input) {
            result.add("'" + s + "'");
        }

        return result.toString().replace("[", "(").replace("]", ")");
    }

    public static Date[] getTimespan(Map<String, Object> propMap, Date reportDate, String timeZone) throws ParseException {
        return getTimespan(propMap, "dateFrom", "dateTo", reportDate, timeZone);
    }

    /** Method used to determine a date interval according to the configured values.
     * This method can handle different property names for dateFrom/dateTo property.
     * Whilst the earlier existing method (which is used very often) {@link #getTimespan(Map, Date, String)} uses 
     * <b>dateFrom</b> as property name for the from property
     * and <b>dateTo</b> as property name for the to property!!! 
     * 
     * @param propMap
     * @param dateFromPropName
     * @param dateToPropName
     * @param reportDate
     * @param timeZone
     * @return
     * @throws ParseException
     */
    public static Date[] getTimespan(Map<String, Object> propMap, String dateFromPropName, String dateToPropName,
    		Date reportDate, String timeZone) throws ParseException {
        Date[] result = new Date[2];

        for (Entry<String, Object> entry : propMap.entrySet()) {
            if (entry.getValue() instanceof String) {
                String value = (String) entry.getValue();

                if (dateFromPropName.equals(entry.getKey())) {
                    if (value.contains("REPORTDATE")) {
                        result[0] = new Date(Util.convertString(value, reportDate, timeZone).getTime());
                    } else {
                        result[0] = DataFormatter.getInstance().getDateFormat(DataFormatter.DATE_TIME_FORMAT, timeZone).parse(value);
                    }
                }

                if (dateToPropName.equals(entry.getKey())) {
                    if (value.contains("REPORTDATE")) {
                        result[1] = new Date(Util.convertString(value, reportDate, timeZone).getTime());
                    } else {
                        result[1] = DataFormatter.getInstance().getDateFormat(DataFormatter.DATE_TIME_FORMAT, timeZone).parse(value);
                    }
                }
            }
        }

        return result;
    }

    /*
    public static boolean validate(ReportdistributorDocument doc) {
        // Set up the validation error listener.
        ArrayList<Object> validationErrors = new ArrayList<Object>();
        XmlOptions validationOptions = new XmlOptions();
        validationOptions.setErrorListener(validationErrors);
        validationOptions.setLoadLineNumbers();

        // During validation, errors are added to the ArrayList for
        // retrieval and printing by the printErrors method.
        boolean isValid = doc.validate(validationOptions);

        // Print the errors if the XML is invalid.
        if (!isValid) {
            Iterator<?> iter = validationErrors.iterator();
            while (iter.hasNext()) {
                XmlError error = (XmlError) iter.next();
                LOG.debug(">> " + error.toString() + " Line:" + error.getLine() + " Column:" + error.getColumn() + "\n");
            }
        }

        return isValid;
    }
    */

    public static boolean getBooleanParam(String paramName, Map<String, Object> propertyMap) {
        return getBooleanParam(paramName, propertyMap, false);
    }

    public static boolean getBooleanParam(String paramName, Map<String, Object> propertyMap, boolean defaultValue) {
        boolean result = defaultValue;
        String valueStr = (String) propertyMap.get(paramName);
        if (valueStr != null && !valueStr.isEmpty() && "true".equals(valueStr)) {
            result = true;
        }
        if (valueStr != null && !valueStr.isEmpty() && "false".equals(valueStr)) {
            result = false;
        }

        return result;
    }

    public static short getShortParam(String paramName, Map<String, Object> propertyMap) {
        short result = (short) -1;
        String valueStr = (String) propertyMap.get(paramName);
        if (valueStr != null && !valueStr.isEmpty()) {
            Short value = Short.parseShort(valueStr);
            if (value != null) {
                result = value.shortValue();
            }
        }

        return result;
    }

    public static int getIntParam(String paramName, Map<String, Object> propertyMap) {
        int result = -1;
        String valueStr = (String) propertyMap.get(paramName);
        if (valueStr != null && !valueStr.isEmpty()) {
            Integer value = Integer.parseInt(valueStr);
            if (value != null) {
                result = value.intValue();
            }
        }

        return result;
    }

    public static String getPlaceHolders(int numValues) {
        StringBuilder result = new StringBuilder();
        for (int i = 0; i < numValues; i++) {
            result.append("?, ");
        }
        result.delete(result.length() - 2, result.length());

        return result.toString();
    }

    public static void setParametersInt(Collection<Integer> values, int startIndex, PreparedStatement stmt) throws SQLException {
        Iterator<Integer> iter = values.iterator();
        int i = startIndex;
        while (iter.hasNext()) {
            stmt.setInt(i++, iter.next());
        }
    }

    public static void setParametersShort(Collection<Short> values, int startIndex, PreparedStatement stmt) throws SQLException {
        Iterator<Short> iter = values.iterator();
        int i = startIndex;
        while (iter.hasNext()) {
            stmt.setShort(i++, iter.next());
        }
    }

    public static void setParametersLong(Collection<Long> values, int startIndex, PreparedStatement stmt) throws SQLException {
        Iterator<Long> iter = values.iterator();
        int i = startIndex;
        while (iter.hasNext()) {
            stmt.setLong(i++, iter.next());
        }
    }

    public static void setParametersDouble(Collection<Double> values, int startIndex, PreparedStatement stmt) throws SQLException {
        Iterator<Double> iter = values.iterator();
        int i = startIndex;
        while (iter.hasNext()) {
            stmt.setDouble(i++, iter.next());
        }
    }

    public static void setParametersFloat(Collection<Float> values, int startIndex, PreparedStatement stmt) throws SQLException {
        Iterator<Float> iter = values.iterator();
        int i = startIndex;
        while (iter.hasNext()) {
            stmt.setFloat(i++, iter.next());
        }
    }

    public static void setParametersString(Collection<String> values, int startIndex, PreparedStatement stmt) throws SQLException {
        Iterator<String> iter = values.iterator();
        int i = startIndex;
        while (iter.hasNext()) {
            stmt.setString(i++, iter.next());
        }
    }

     
    

    public static boolean isDateBetween(Date date, Date start, Date end) {
        return !date.before(start) && ! date.after(end);
    }
    
    public static String addLinebreaksToString (String string, int addLinebreakAfterCharactersAmount)
    {
    	if (string == null || string.length() == 0)
    	{
    		return string;
    	}
    	
    	StringBuffer resultBuf = new StringBuffer();
    	
    	int startIndex = 0;
    	int endIndex = addLinebreakAfterCharactersAmount;
    	char charAt;
    	int i;
    	
    	String subString;
    	while (true)
    	{
    		if (endIndex > string.length())
    		{
    			endIndex = string.length();
    		}
    		
    		subString = string.substring(startIndex, endIndex);
    		
    		//don't cut in the middle of the word, only break if the last character is a blank
    		for (i=endIndex;i<string.length();i++)
    		{
    			charAt = string.charAt(i);
    			if (String.valueOf(' ').equals(String.valueOf(charAt)))
    			{
    				break;
    			}
    			subString = string.substring(startIndex, i);
    		}
    		
    		subString = string.substring(startIndex, i).trim();
    		resultBuf.append(subString).append("\n");
    		startIndex = i;
    		endIndex += subString.length();
    		
    		if (startIndex >= string.length())
    		{
    			break;
    		}
    	}
    	
    	return resultBuf.toString();
    }
    
    /**
     * Parameter should be formed like 134,345,5676... 
     * @param propertyMap
     * @param paramName
     * @return Collection of Integer
     */
    public static Collection<Integer> getIntegerCollectionByParam(Map<String, Object> propertyMap, String paramName) {
    	
        Collection<Integer> result = null;
        String valueStr = (String) propertyMap.get(paramName);
        if (valueStr != null && !valueStr.isEmpty()) {
        	
        	result = new ArrayList<Integer>();
        	String[] aValues =  valueStr.split("[,]");
        	
        	for (String _value : aValues) {
        		Integer value = Integer.parseInt(_value.trim());
        		if (value != null) {
                    Integer id = value.intValue();
                    result.add(id);
                }
			}
        }

        return result;
    }

    /**
     * Parameter should be formed like 123,456,7891...
     * @param propertyMap
     * @param paramName
     * @return Set of Short
     */
    public static Set<Short> getShortSetByParam(Map<String, Object> propertyMap, String paramName) {

        Set<Short> result = null;
        String valueStr = (String) propertyMap.get(paramName);
        if (valueStr != null && !valueStr.isEmpty()) {

            result = new HashSet<>();
            String[] aValues =  valueStr.split("[,]");

            for (String _value : aValues) {
                Short value = Short.parseShort(_value.trim());
                if (value != null) {
                    Short id = value.shortValue();
                    result.add(id);
                }
            }
        }

        return result;
    }
    
    
    /**
     * Convert Collection of Integer to String like 123,345,345  or '123','345'...
     * @param arg
     * @param apostrophe
     * @return String
     */
    
    public static String convertIntegerCollectionToString(Collection<Integer> arg, String apostrophe) {
    	String result = "";
    	StringBuffer sb = new StringBuffer();
    	for (Integer _id : arg) {
    		if (sb.length() > 0) {
    			sb.append(",");
    		}
    		sb.append(apostrophe);
    		sb.append(_id);
    		sb.append(apostrophe);
		}
    	result = sb.toString(); 
    	return result;
    }
    
    public static Timestamp convertDate2Timestamp (Date date)
    {
    	Timestamp timestamp = null;
    	if (date!=null)
    	{
    		timestamp = new Timestamp(date.getTime());
    	}
    	
    	return timestamp;
    }
    
    public static Date convertTimestamp2Date (Timestamp timestamp)
    {
    	Date date = null;
    	if (timestamp!=null)
    	{
    		date = new Date (timestamp.getTime());
    	}
    	
    	return date;
    }
    
    public static List<Integer> convert2IntList(String arg) {
    	List<Integer> result = new ArrayList<>();
    	String[] aElems = new String[]{};
    	boolean splitable = false;
    	if (arg.indexOf(";") > -1) {
    		aElems = arg.split("[;]");
    		splitable = true;
    	}
    	if (arg.indexOf(",") > -1) {
    		aElems = arg.split("[,]");
    		splitable = true;
    	}
    	
    	if (splitable) {
	    	for (String _elem : aElems) {
				Integer id = Integer.parseInt(_elem);
				result.add(id);
			}
    	} else {
    		Integer id = Integer.parseInt(arg);
    		result.add(id);
    	}
    	
    	return result;

	}
    
    public static int getMonthOfDate(Date date) {
    	Calendar cal = Calendar.getInstance();
        cal.setTime(date);
        int _calmonth = cal.get(Calendar.MONTH);
        return _calmonth;
    }

    public static int getYearOfDate(Date date) {
        Calendar cal = Calendar.getInstance();
        cal.setTime(date);
        int _calmonth = cal.get(Calendar.YEAR);
        return _calmonth;
    }

    public static String getNameOfMonthFromInt(int num) {
        String month = "wrong";
        Locale loc = Locale.ENGLISH;
        DateFormatSymbols dfs = new DateFormatSymbols(loc);
        String[] months = dfs.getMonths();
        if (num >= 0 && num <= 11 ) {
            month = months[num];
        }
        return month;
    }

    private static void changeDateTimePart(Calendar cal, String val, DateFunction fun) {
        int part = Calendar.DAY_OF_YEAR;
        String value = val;
        int len = value.length();

        if (len > 1) {
            char c = value.charAt(len-1);
            switch (c) {
                case 'Y':
                    part = Calendar.YEAR;
                    value = value.substring(0, value.length()-1);
                    break;
                case 'M':
                    part = Calendar.MONTH;
                    value = value.substring(0, value.length()-1);
                    break;
                case 'W':
                    part = Calendar.WEEK_OF_MONTH;
                    value = value.substring(0, value.length()-1);
                    break;
                case 'D':
                    part = Calendar.DAY_OF_MONTH;
                    value = value.substring(0, value.length()-1);
                    break;
                case 'h':
                    part = Calendar.HOUR;
                    value = value.substring(0, value.length()-1);
                    break;
                case 'm':
                    part = Calendar.MINUTE;
                    value = value.substring(0, value.length()-1);
                    break;
                case 's':
                    part = Calendar.SECOND;
                    value = value.substring(0, value.length()-1);
                    break;
                default:
                    LOG.warn("Unsupported value [" + c + "]");	
            }
        }

        switch(fun) {
            case ADD:
                cal.add(part, Integer.valueOf(value));
                break;
            case SUB:
                cal.add(part, Integer.valueOf(value) * -1);
                break;
        }
    }
    
    public static boolean isOracleConnection(Connection con) {
    	return (con != null && con.toString().contains("oracle"));
    }

    public static boolean isOracleConnection(ComboPooledDataSource con) {
        return (con != null && con.getJdbcUrl().contains("oracle"));
    }
 
	public static Date unsetMilliseconds(Date date)
	{
		Calendar cal = Calendar.getInstance();
		cal.set(Calendar.MILLISECOND, 0);
		return cal.getTime();
	}
    
    public static boolean valueIsDateTime(String par) {
		boolean back = false;
		
		boolean isTime = false;
		if (par != null) {
			String[] aElems =  par.split(":");
			if (aElems.length == 3) {
				isTime = true;
			}
		}
		
		back = isTime;
		
		return back;
	}

    
    /**
     * Parameter should be formed like 'foo,bar,barfoo' 
     * @param propertyMap
     * @param paramName
     * @return Collection of String
     */
    public static Collection<String> getStringCollectionByParam(Map<String, Object> propertyMap, String paramName) {
    	
        Collection<String> result = null;
        String valueStr = (String) propertyMap.get(paramName);
        if (valueStr != null && !valueStr.isEmpty()) {
        	
        	result = new ArrayList<>();
        	String[] aValues =  valueStr.split("[,]");
        	
        	for (String _value : aValues) {
                result.add(_value.trim());
			}
        }

        return result;
    }
	
	public static Date convertGmtDateToLocalTime (Timestamp timestamp, String timezone)
	{
		TimeZone dtz = TimeZone.getTimeZone(timezone);
		long diffMillSec = dtz.getOffset(timestamp.getTime());
		long localTimeMillis = timestamp.getTime() + diffMillSec;
		return new Date (localTimeMillis);
	}
	

	

    public static Boolean getOptionalBooleanValue(ResultSet rs, int rsIdx) throws SQLException
    {
        Boolean ret = null;

        if (rs.getObject(rsIdx)!=null)
        {
            ret = rs.getBoolean(rsIdx);
        }
        return ret;
    }

    public static Integer getOptionalIntegerValue(ResultSet rs, int rsIdx) throws SQLException
    {
        Integer ret = null;

        if (rs.getObject(rsIdx)!=null)
        {
            ret = rs.getInt(rsIdx);
        }
        return ret;
    }

    public static String getFormattedTime(int secondsIn) {
        final DecimalFormat DF2 = new DecimalFormat("00");
        final DecimalFormat DF5 = new DecimalFormat("###00");

        int seconds = secondsIn % 60;
        int value = secondsIn / 60;
        int minutes = value % 60;
        value = value / 60;
        int hours = value % 60;
        String result = DF5.format(hours) + ":" + DF2.format(minutes) + ":" + DF2.format(seconds);

        while (result.contains("-")) {
            result = result.replaceAll("-", "");
        }

        if (secondsIn < 0) {
            result = "-" + result;
        }

        return result;
    }
    
    public static Collection<Short> convertStringToShortCollection (String string)
    {
    	return convertStringToShortCollection(string, ",");
    }
    
    public static Collection<Short> convertStringToShortCollection (String string,
    		String separator)
    {
    	Collection<Short> result = new ArrayList<>();
    	
    	for (String s : string.split(separator))
    	{
    		if (! string.isEmpty())
    		{
    			result.add(Short.valueOf(s));
    		}
    	}
    	
    	return result;
    }
    
    public static Integer convert2Int (String string)
    {
    	Integer result = null;
    	if (! string.isEmpty())
    	{
    		result = Integer.valueOf(string);
    	}
    	return result;
    }
    
    public static Boolean convert2Boolean (String string)
    {
    	Boolean result = null;
    	if (! string.isEmpty())
    	{
    		result = Boolean.valueOf(string);
    	}
    	return result;
    }
    
    public static Short convert2Short (String string)
    {
    	Short result = null;
    	if (! string.isEmpty())
    	{
    		result = Short.valueOf(string);
    	}
    	return result;
    }

	/**
	 * Happens often that event code array should be converted to set
	 * @param eventCodes Events codes added to set
	 * @return set with given event codes, empty if given list is null
	 */
	public static Set<Short> convertToSet(short ... eventCodes) {
		Set<Short> result = new LinkedHashSet<Short>();
		if (eventCodes!=null) {
			for (short eventCode : eventCodes) {
				result.add(eventCode);
			}
		}
		return result;
	}



  public static boolean isMySqlConnection(Connection con) {
    return (con != null && con.toString().contains("mysql"));
  }

  public static boolean isMySqlConnection(ComboPooledDataSource con) {
    return (con != null && con.getJdbcUrl().contains("mysql"));
  }

  /**
   * Extracted from DataManager / GenericReport / GenericExtendedReport ... readMetaData
   * @param itemRecord
   * @param typeName
   * @param precision
   * @param scale
   * @param paramItem
   */
  public static void addItemType(
      ItemRecord itemRecord,
      String typeName,
      int precision,
      int scale,
      ItemRecord paramItem) {

	  typeName = typeName.toLowerCase(); // all check for types should ignore case !

    if (typeName.indexOf("int") > -1) {
      itemRecord.addItem(ITEMTYPE, ItemRecord.INTEGER, ItemRecord.STRING);
    } else if (typeName.indexOf("datetime") > -1) {
      itemRecord.addItem(ITEMTYPE, ItemRecord.DATE2, ItemRecord.STRING);
    } else if (typeName.indexOf("timestamp") > -1) {
      itemRecord.addItem(ITEMTYPE, ItemRecord.DATE2, ItemRecord.STRING);
    } else if (typeName.indexOf("char") > -1) {
      itemRecord.addItem(ITEMTYPE, ItemRecord.STRING, ItemRecord.STRING);
    }  else if (typeName.indexOf("text") > -1) {
      itemRecord.addItem(ITEMTYPE, ItemRecord.STRING, ItemRecord.STRING);
    } else if (typeName.indexOf("bit") > -1) {
      itemRecord.addItem(ITEMTYPE, ItemRecord.BOOLEAN, ItemRecord.STRING);
    } else if (typeName.indexOf("bool") > -1) {
      itemRecord.addItem(ITEMTYPE, ItemRecord.BOOLEAN, ItemRecord.STRING);
    } else if (typeName.indexOf("numeric") > -1) {
      itemRecord.addItem(ITEMTYPE, ItemRecord.DOUBLE, ItemRecord.STRING);
    } else if (typeName.indexOf("decimal") > -1) {
      if (scale == 0) {
        itemRecord.addItem(ITEMTYPE, ItemRecord.INTEGER, ItemRecord.STRING);
      } else {
        itemRecord.addItem("ITEMTYPE", ItemRecord.DOUBLE, ItemRecord.STRING);
      }
    } else if (typeName.indexOf("float") > -1) {
      itemRecord.addItem(ITEMTYPE, ItemRecord.DOUBLE, ItemRecord.STRING);
    } else if (typeName.indexOf("double") > -1) {
      itemRecord.addItem(ITEMTYPE, ItemRecord.DOUBLE, ItemRecord.STRING);
    } else if (typeName.indexOf("number") > -1) {
      if (scale == 0) {
        itemRecord.addItem(ITEMTYPE, ItemRecord.INTEGER, ItemRecord.STRING);
      } else {
        itemRecord.addItem("ITEMTYPE", ItemRecord.DOUBLE, ItemRecord.STRING);
      }
    } else if (typeName.indexOf("lob") > -1) {
      itemRecord.addItem(ITEMTYPE, ItemRecord.STRING, ItemRecord.STRING);
    } else {
      LOG.debug(">>> not implemented for type: " + typeName);
    }

    // type Converter Oracle

    itemRecord.addItem(
        ITEMTYPE,
        readConvertFieldType(
            itemRecord.getString(NAME),
            itemRecord.getString(ITEMTYPE),
            paramItem),
        ItemRecord.STRING);
  }

  /**
   * [fieldName]:Double,Integer
   * @param fieldName
   * @param orgType
   * @param paramItem
   * @return String
   */
  private static String readConvertFieldType(String fieldName, String orgType, ItemRecord paramItem) {
    String result = orgType;
    String[] fields = paramItem.getString(Cf.PAR_FIELD_READ_TYPE).split("[;]");
    for (String valuePears: fields) {
      String[] elems = valuePears.split("[:]");
      if (elems.length > 1) {
        //if (fieldName.equals(oracleString(elems[0]))) {
        if (fieldName.toUpperCase().equals(elems[0].toUpperCase())) {
          result = elems[1];
          break;
        }
      }
    }
    return result;
  }

  /**
   * Retrieve column label (alias) or column name if no alias is set
   * @param metaData
   * @param idx
   * @return
   */
  public static String getColumnName(ResultSetMetaData metaData, int idx) throws SQLException {
    String result = metaData.getColumnLabel(idx);
    if (! result.isEmpty()) {
      result = metaData.getColumnLabel(idx);
    }
    return result;
  }

  public static List<Short> convert2ShortList(String arg) {
    List<Short> result = new ArrayList<>();

    String[] tokens = null;
    if (arg.contains(";")) {
      tokens = arg.split("[;]");
    } else if (arg.contains(",")) {
      tokens = arg.split("[,]");
    }

    if (tokens != null) {
      for (String token : tokens) {
        result.add(Short.valueOf(token));
      }
    } else {
      result.add(Short.valueOf(arg));
    }

    return result;
  }

  public static List<String> convert2StringList(String arg) {
    List<String> result = new ArrayList<>();

    String[] tokens = null;
    if (arg.contains(";")) {
      tokens = arg.split("[;]");
    } else if (arg.contains(",")) {
      tokens = arg.split("[,]");
    }

    if (tokens != null) {
      result.addAll(Arrays.asList(tokens));
    } else {
      result.add(arg);
    }

    return result;
  }

  public static long calculateTimeDifference(Date date1, Date date2) {
	  return (date1.getTime() - date2.getTime()) / 1000;
  }
}
