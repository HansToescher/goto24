package at.goto24.data.util;

import java.text.DecimalFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TimeZone;



public final class FilenameConverter {
    private static final DecimalFormat DF = new DecimalFormat("00");

    private FilenameConverter() {
    }

    public static String convertDateInFileName(String filename) {
	Calendar cal = Calendar.getInstance();
	cal.setTimeInMillis(System.currentTimeMillis());
	return convertDateInFileName(filename, cal);
    }

    public static String convertDateInFileName(String filename, Date date) {
    Calendar cl = Calendar.getInstance();	
    cl.setTime(date);	
	return convertDateInFileName(filename, cl);
    }

    @SuppressWarnings("unchecked")
    public static String convertDateInFileName(String parFileName, String timezone, Calendar cal) {
    	cal.setTimeZone(TimeZone.getTimeZone(timezone));

	    return convertDateInFileName(parFileName, cal);
    }

    private static Integer getValue(Map<String, Integer> map, String key) {
	return (map.get(key) == null ? Integer.valueOf(0) : map.get(key));
    }

    public static String convertDateInFileName(String filename, Calendar cal) {
	String result = filename;
	if (result.contains("%")) {
	    result = result.replaceAll("%y", String.valueOf(cal.get(Calendar.YEAR)));
	    result = result.replaceAll("%M", DF.format(cal.get(Calendar.MONTH) + 1));
	    result = result.replaceAll("%d", DF.format(cal.get(Calendar.DAY_OF_MONTH)));
	    result = result.replaceAll("%H", DF.format(cal.get(Calendar.HOUR_OF_DAY)));
	    result = result.replaceAll("%m", DF.format(cal.get(Calendar.MINUTE)));
	    result = result.replaceAll("%s", DF.format(cal.get(Calendar.SECOND)));
	}
	return result;
    }
    
    public static String replacePlaceholdersInFilename (String curFilename,
    		Map<String, String> placeholderReplacementMap)
	{
    	String resultFilename = curFilename;
    	
    	if (placeholderReplacementMap!=null)
    	{
	    	for (Entry<String, String> mapEntry : placeholderReplacementMap.entrySet())
	    	{
	    		String propName = mapEntry.getKey();
	    		String propValue = mapEntry.getValue();
	    		
	    		resultFilename = resultFilename.replace(propName, propValue);
	    	}
    	}
    	
    	return resultFilename;
	}
}
