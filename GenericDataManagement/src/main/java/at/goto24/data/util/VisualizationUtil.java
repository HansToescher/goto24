package at.goto24.data.util;

import org.apache.commons.collections4.*;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;

import java.io.File;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.util.*;


public class VisualizationUtil {
	private static final Logger LOG = Logger.getLogger(VisualizationUtil.class);
	private static final File baseDir = new File(System.getProperty("user.dir") + File.separator + "output" + File.separator + "visualizationreport");

	public enum StatType {
		VERSION,
		APP_TYPE,
		CALLING_TYPE
	}

	public static File getBaseDir() {
		return baseDir;
	}

	public static Collection<Date[]> splitByDays(Date[] timeSpan, String timezone) {
		if (null == timeSpan || 2 != timeSpan.length) {
			LOG.error("Time span was not specified");
			return Collections.emptyList();
		}

		Date from = timeSpan[0];
		Date to = timeSpan[1];
		ZonedDateTime fromDate = from.toInstant().atZone(ZoneId.of(timezone));
		ZonedDateTime toDate = to.toInstant().atZone(ZoneId.of(timezone));
		Duration duration = Duration.between(fromDate, toDate);
		long days = duration.toDays();
		Collection<Date[]> ranges = new ArrayList<>();
		if (0 != days) { // more than 1 day
			for (int i = 0; i <= days; i++) {
				ZonedDateTime nextEnd = fromDate.plusDays(1).minusSeconds(1);
				if (nextEnd.isAfter(toDate)) {
					nextEnd = toDate;
				}
				Date newFrom = new Date(fromDate.toEpochSecond() * 1000);
				Date newTo = new Date(nextEnd.toEpochSecond() * 1000);
				ranges.add(new Date[]{newFrom, newTo});
				fromDate = fromDate.plusDays(1);
			}
		} else {
			ranges.add(new Date[]{from, to});
		}

		if (CollectionUtils.isNotEmpty(ranges)) {
			LOG.info("Report will look into:");
			ranges.forEach(dates -> LOG.info(String.format("from [%s] to [%s]", dates[0], dates[1])));
		}

		return ranges;
	}


	public static String formatAvgDuration(double avgDurationSeconds) {
		SimpleDateFormat dateFormat = new SimpleDateFormat("HH:mm:ss.SSS");
		dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
		return dateFormat.format(new Date((int) (avgDurationSeconds * 1000)));
	}

    public static void trackAllColumnsForAutoSizing(Sheet sheet) {
        if (null != sheet && sheet instanceof SXSSFSheet) {
            SXSSFSheet xSheet = (SXSSFSheet) sheet;
            xSheet.trackAllColumnsForAutoSizing();
        }
    }
}
