package at.goto24.data.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashSet;
import java.util.Locale;
import java.util.Set;
import java.util.Map.Entry;

import at.goto24.data.ItemRecord;

/**
 * Tools for Flexible Reports
 * @author j.toescher
 * @version 1.0.0
 * @since 29.09.2016
 */
public class Tools {
	
	public static String getTitle(String parTitle, ItemRecord _params) {
		String result = parTitle;
		
		for (Entry<String, ItemRecord> _i : _params.getMapItems().entrySet()) {
			String param = _i.getKey();
			String search = "\\$"+param;
			result = result.replaceAll(search, _params.getString(param));
		}   
		
		return result;
	}
	
	
	/**
	 * Overload
	 * @param part
	 * @param millSec
	 * @return
	 */
	public static String partOfDate(String part, Long millSec) {
		Date date = new Date(millSec);
		return partOfDate(part, date);
	}
	
	/**
	 * 
	 * @param part 'm' Monat, 'y' Year, 'd' Day 
	 * @param date
	 * @return String '2016' or '01..12' '01..31'
	 */
	public static String partOfDate(String part, Date date) {
		String result = "";
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
		String _date = sdf.format(date);
		
		if ("m".equals(part)) { // Monat  -> 01..12
			result = _date.substring(5, 7);
		} else if ("y".equals(part)) { // Year -> 2010..2016
			result = _date.substring(0, 4);
		} else if ("d".equals(part)) { // Day -> 01..31
			result = _date.substring(8, 10);
		} else if ("date".equals(part)) { // Day -> 01..31
			result = _date.substring(0, 10);
		} else if ("time".equals(part)) { // Day -> 01..31
			result = _date.substring(11, 19);
		} else if ("datetime".equals(part)) { // String
			result = _date;
		}
		
		return result;
	}
	
	
	public static Date getDate(String argDate) throws ParseException {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
		return sdf.parse(argDate);
	}
	
	/**
	 * 
	 * @param argDate String
	 * @param datePattern eg. 'yyyy-MM-dd HH:mm:ss.SSS' 
	 * @return type of Date()
	 * @throws ParseException
	 */
	public static Date getDate(String argDate, String datePattern) throws ParseException {
		SimpleDateFormat sdf = new SimpleDateFormat(datePattern);
		return sdf.parse(argDate);
	}
	
	
	/**
	 * Convert date String to Date type.
	 * @param argDate String
	 * @param datePattern String
	 * @param locale Locale
	 * @return Date
	 * @throws ParseException
	 */
	public static Date getDate(String argDate, String datePattern, Locale locale) throws ParseException {
		SimpleDateFormat sdf = new SimpleDateFormat(datePattern, locale);
		return sdf.parse(argDate);
	}
	
	/**
	 * Convert date to formated string value
	 * @param argDate
	 * @param datePattern
	 * @param locale
	 * @return String
	 * @throws ParseException
	 */
	public static String getDateFormated(Date argDate, String datePattern, Locale locale) throws ParseException {
		SimpleDateFormat sdf = null;
		if (locale != null) {
			sdf = new SimpleDateFormat(datePattern, locale);
		} else {
			sdf = new SimpleDateFormat(datePattern);
		}
		return sdf.format(argDate);
	}
	
	
	public static Date getDateFrom(String year, String month) throws ParseException {
		Date result = null;
		
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
		String xmon = month.length() == 1 ? "0"+month : month;
		String xDate = year+"-"+xmon+"-01 00:00:00.0";
		result = sdf.parse(xDate);
		
		return result;
	}
	
	public static String maxDay(String month, String year) {
		String result = "31";
		if ("02".indexOf(month) >= 0) {
			if ((Integer.valueOf(year) % 4) == 0 ) {
				result = "29";
			} else {
				result = "28";
			}
		}
		if ("04,06,09,11".indexOf(month) >= 0) {
			result = "30";
		}
		
		return result;
	}

	
	public static Date getDateTo(String year, String month) throws ParseException {
		Date result = null;
		
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
		String xmon = month.length() == 1 ? "0"+month : month;
		String xday = maxDay(xmon,year);
		String xDate = year + "-" + xmon + "-" + xday+" 23:59:59.900";
		result = sdf.parse(xDate);
		
		return result;
	}
	
	
	public static Set<Integer> getSetInteger(String par) {
		Set<Integer> result = new HashSet<Integer>();
		String[] ids = par.split(",");
		for (String id : ids) {
		    Integer _id = Integer.parseInt(id);
		    result.add(_id);
		}
		return result;
	}
	
	public static Collection<Integer> getListOfInteger(String par) {
		Collection<Integer> result = new ArrayList<Integer>();
		if (par != null && !"".equals(par)) {
			String[] elems = par.split("[,]");
			for (String elem : elems) {
				Integer id = Integer.parseInt(elem);
				result.add(id);
			}
		}
		
		return result;
	}
	
	public static String convert2LowUpperCases(String argString, String caseType) {
		String back = argString;
		
		if ("XX".equals(caseType)) {
			back = argString.toUpperCase();
		} else if ("xx".equals(caseType)) {
			back = argString.toLowerCase();
		} else if ("Xx".equals(caseType)) {
			StringBuffer sb = new StringBuffer();
			int x = 0;
			for (char _x :argString.toCharArray()) {
				
				if (x == 0) {
					String _l = new StringBuffer().append(_x).toString().toUpperCase();
					sb.append(_l);
				} else {
					String _l = new StringBuffer().append(_x).toString().toLowerCase();
					sb.append(_l);
				}
				
				x++;
				
				if (_x == ' ') {
					x = 0;
				}
			}
			back = sb.toString();
		}
		
		return back;
	}
	
	/**
	 * Convert a String value formated as 'HH:mm:ss' or 'HH:mm' to target values as String.<br>
	 * 
	 * @param argTarget possible values 'H' Hours, 'M' Minutes, 'S' Seconds.
	 * @param argBaseTime as String formated like 'HH:mm:ss' or 'HH:mm'
	 * @return targetValue as String 
	 */
	public static String convertTimeString_to_X(String argTarget, String arg_BaseTime, String faultValue) {
		String result = faultValue;
		
		try {
		String argBaseTime = arg_BaseTime.trim();
		if ("H".equals(argTarget)) { // Hours
			String[] elems = argBaseTime.split("[:]");
			if (elems.length == 3) {
				String xh = elems[0];
				String xm = elems[1];
				String xs = elems[2];
				
				Integer _h = Integer.parseInt(xh);
				
				Integer _xm = Integer.parseInt(xm);
				Integer _xs = Integer.parseInt(xs);
				if (_xs.intValue() > 30) {
					_xm = _xm + 1;
				}
				if (_xm.intValue() > 30) {
					_h = _h + 1;
				}
				
				result = "" + _h;
			}
			
			if (elems.length == 2) {
				String xh = elems[0];
				String xm = elems[1];
				Integer _h = Integer.parseInt(xh);
				
				Integer _xm = Integer.parseInt(xm);
				if (_xm.intValue() > 30) {
					_h = _h + 1;
				}
				
				result = "" + _h;

			}
			
		}
		
		if ("M".equals(argTarget)) { // Minutes
			String[] elems = argBaseTime.split("[:]");
			if (elems.length == 3) {
				String xh = elems[0];
				String xm = elems[1];
				String xs = elems[2];
				Integer _m = Integer.parseInt(xm) + (60*Integer.parseInt(xh));
				
				Integer _s = Integer.parseInt(xs);
				if (_s.intValue() > 30) {
					_m = _m + 1;
				}
				
				result = "" + _m;
			}
			
			
			if (elems.length == 2) {
				String xh = elems[0];
				String xm = elems[1];
				
				Integer _m = Integer.parseInt(xm) + (60*Integer.parseInt(xh));
				result = "" + _m;
			}

			
		}
		
		if ("S".equals(argTarget)) { // Seconds
			String[] elems = argBaseTime.split("[:]");
			if (elems.length == 3) {
				String xh = elems[0];
				String xm = elems[1];
				String xs = elems[2];
				
				Integer _s = (60 * Integer.parseInt(xm)) + Integer.parseInt(xs) + (60*60*Integer.parseInt(xh)); 
				
				result = "" + _s;
			}
			
			if (elems.length == 2) {
				String xh = elems[0];
				String xm = elems[1];
				
				Integer _s = (60 * Integer.parseInt(xm))  + (60*60*Integer.parseInt(xh)); 
				
				result = "" + _s;
			}
		}
		
		} catch (Exception ex) {
			result = faultValue;
		}
		
		return result;
	}
	
	
}
