package at.goto24.data.util;

import java.awt.Color;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public final class HSSFUtil {
    private static final Logger LOG = Logger.getLogger(HSSFUtil.class);

    public enum StyleType {
	style, styleYellow, styleRed, styleBold, titleGrey, styleYellowLeftBorder, styleBoldCenter,

	styleLeftBorder, styleBoldPercent, styleBoldNoBorder, stylePercent, stylePercentBorder, stylePercentBorderGrey, stylePercentRightMediumBorderNoDetails,

	styleNumber, styleNumberBorder, stylePercentDetailed, styleYellowBlackBoldCenter, styleCentered,

	styleYellowBlackBoldRight, styleOrangeHeader, styleOrangeHeaderRight, styleBoldLeftBoldBorder, styleGreyBoldLeftRightBoldBorder,

	styleBottomBorder, styleRightBoldBorderBottomThin, styleOrangeFooter, styleBoldLeftBottomBoldBorder,

	styleBottomBoldBorder, styleBoldBottom, styleBoldBottomRight, styleBoldGreyLeftBold, styleBoldGrey,

	styleBoldGreyRightBold, styleLeftBoldBorder, styleRightBoldBorder, styleRightBoldBorderBottomThinRed,

	styleBoldBottomRightRed, styleOrangeFooterRed, styleRightBoldBorderRed, styleLeftBottomBoldBorder,

	styleBoldLeftBoldBorderRedBg, styleRedBg, styleRightBoldBorderBottomThinRedBg,

	styleBoldLeftBoldBorderYellowBg, styleYellowBg, styleRightBoldBorderBottomThinYellowBg,

	styleNumberDetailed, styleNumberDetailedBold, styleNoBorder, styleLeftTopRightBoldBorder,

	styleLeftTopBoldBorder, styleTopRightBoldBorder, styleLeftBottomRightBoldBorder,

	styleLeftBottomBoldBorder2, styleBottomRightBoldBorder, styleLeftRightBoldBorder,

	styleWhiteBlackBoldHeader, styleWhiteBlackBoldHeaderCenter, styleWhiteBlackBoldNumberRight, styleWhiteBlackBoldPercentRight, styleBoldTopRightMedium, styleRightMedium, styleBoldBottomRightMedium,

	styleLeftMedium, styleAllThin, styleBoldBottomLeftMedium, styleBoldAllMedium, styleBoldBottomMedium, styleBoldAllThin,

	styleBoldLeft, styleBoldRight, styleBottomLeftMedium, styleBottomMedium, styleBottomRightMedium,

	blue, white, blueNumber, whiteNumber, styleDate, styleTime, styleDate2, styleTime2, styleTimeBorder, styleTime2NoBorder, styleDateGermanStandardCentered,

	styleNumberDetailedBorder, styleNumberBorderNoDetails, styleNumberBorderMediumLeftNoDetails, styleDateTime,

	styleDateTimeOrange, styleDateTimeRed, styleTopLeftBottom, styleTopBottom, styleTopRightBottom,

	styleRightLeft, styleWhiteBlackBoldHeaderCenterCalibri, styleWhiteBlackBoldHeaderLeftCalibri,

	styleGreyBoldLeftRightBoldBorderCalibri, styleGreyBoldRightBoldBorderCalibri,

	styleWhiteBlackBoldHeaderRightCalibri, styleGreyBoldRightLeftRightBoldBorderCalibri, styleDateTimeBorder,

	styleAllMediumCalibri, styleNumberBorderRightMedium, styleBottomMediumCalibri,

	styleNumberBorderBottomMediumNoDetails, styleNumberBorderBottomMedium, styleNumberBorderBottomRightMedium,

	styleDateTimeBorderBottomMedium, styleLeftRightBoldBorderRight, styleLeftRightBoldBorderRightRedCalibri,

	styleGreyBoldLeftCalibri, styleGreyRightBorderNoDetailsBold, styleLeftRightBoldBorderRightGrey,

	styleLeftRightBoldBorderRightRedCalibriGrey, styleGreyBoldLeftCalibriBottom, styleGreyRightBottomBorderNoDetailsBold,

	styleLeftBottomBoldBorderRightRedCalibriGrey, styleLeftBoldBorderRightBottomGrey, styleWhiteBlackBoldHeaderLeftCalibriWrap,

	styleWhiteBlackBoldHeaderLeftCalibriTopText, stylePercentNoDetails, stylePercentNoDetailsRightBorder,

	styleGreyBoldRightCalibri, styleGreyBoldRightCalibriNumber, styleLeftRightBoldBorderRightGreyPercent,

	styleGreyBoldRightCalibriBottom, styleGreyBoldRightCalibriNumberBottom, styleLeftRightBoldBorderRightBottomGreyPercent,

	styleYellowRightBorder, styleBoldTopLeftBottomBorder, styleBoldTopRightBottomBorder,

	titleGreyAllMedium, styleBoldTopBottom, styleLeftRightBorder, styleYellowLeftRightBorder,

	styleYellowTopBottomBorder, styleTopBottomBorder, styleYellowLeftBorderRightNoBorder,

	styleYellowRightBorderLeftNoBorder, styleLeftBorderRightNoBorder, styleRightBorderLeftNoBorder,

	styleNoBorderBold, styleMonth, styleNumberBorderMediumRightNoDetails, styleBoldLeftRight,

	styleDateTimeArial, styleArial, styleRedArial, styleNormalArial, stylePercentBorderNoDetails,

	stylePercentBorderBottomNoDetails, styleNumberBorderBottomRightMediumNoDetails,

	styleNumberBoldBorderRightGreyBGNoDetails, styleBoldBorder, styleBoldBorderBottom,

	styleNumberBoldBorderRightBottomGreyBGNoDetails, styleNumberBorderRightMediumNoDetails,

	styleNumberBorderNoDetailsRed, styleNumberBorderBottomMediumNoDetailsRed,

	styleNumberBorderRightMediumNoDetailsRed, styleNumberBorderBottomRightMediumNoDetailsRed,
	
	styleDateTimeWithSeconds, styleBoldCenterLightGreen, styleArialWrapText, styleBottomRightBorder, styleRightBorder,
	
	styleLeftBorderRightBorder, styleLeftBorderRightBorderBottomBorder, stylePercentDetailed1DecimalPlace,
	
	styleDateTimeBorderGreen, styleLeftBorderGreen, styleTimeBorderGreen, styleDateBorder, styleDateBorderGreen, styleDateBorderLightYellow,
	
	styleLeftBorderGrey, styleNumberBorderGrey, styleDateTimeWithSecondsBorder, styleTimeWithSecondsBorder,
	
	styleLeftBorderOrange, styleLeftBorderRed, styleBoldCenterLightYellow, styleBoldCenterLightRed,
	styleNumberDezimal3Border, styleNumberDezimal1Border, styleNumberDezimal4Border, styleNumberDezimal5Border,
	
	styleGreenNoBorder, styleRedNoBorder, styleOrangeBorder, styleDateOrangeBorder, styleTimeBorderLightYellow, styleTimeOrangeBorder,
	
	styleFillColorLightGreenBoldCentered, styleNumberLightGreenBoldCentered, styleNumberRedBoldCentered, styleRoseNoBorderBoldCentered, styleRedArialBoldCentered,styleRightAligned
    };

    public enum FontType {
	bold, orangeBold, whiteBold, blackBold, black, blue, red, whiteBoldCalibri, boldCalibri, calibri, redCalibri, redCalibriBold, green
    };

    private static final Map<Workbook, Map<FontType, Font>> FONTS = new HashMap<Workbook, Map<FontType, Font>>();

    private static final Map<Workbook, Map<StyleType, CellStyle>> STYLES = new HashMap<Workbook, Map<StyleType, CellStyle>>();

    private static final List<String> COLUMN_NAMES = new ArrayList<String>();

    private static final int COLUMN_NAME_LIMIT = 500;

    // important: HSSF and XSSF have a different color handling.
    // * in HSSF the default color palette must be adapted, then the color is
    // used via its palette index
    // * in XSSF the color can be set directly to the object as XSSFColor, the
    // XSSF object has to be cast correctly
    private static final Color COLOR_ORANGE = new Color(255, 192, 0);

    private static final Color COLOR_LIGHT_BLUE = new Color(219, 229, 241);

    private static final Color COLOR_BLUE = new Color(92, 96, 164);

    private static final Color COLOR_GREY_50_PERCENT = new Color(240, 240, 240);

    private static final Color COLOR_RED = new Color(255, 0, 0);

    private static final XSSFColor ORANGE = new XSSFColor(COLOR_ORANGE);

    private static final XSSFColor LIGHT_BLUE = new XSSFColor(COLOR_LIGHT_BLUE);

    private static final XSSFColor BLUE = new XSSFColor(COLOR_BLUE);

    private static final XSSFColor GREY_50_PERCENT = new XSSFColor(COLOR_GREY_50_PERCENT);

    private static final XSSFColor RED = new XSSFColor(COLOR_RED);

    static {
	fillColumnNames();
    }

    private HSSFUtil() {
    }

    private static void fillColumnNames() {
	for (int i = 65; i < 91; i++) {
	    COLUMN_NAMES.add((char) i + "");
	}
	StringBuilder entry = new StringBuilder();
	for (int i = 65; i < 91; i++) {
	    entry.append((char) i);
	    for (int j = 65; j < 91; j++) {
		entry.append((char) j);
		if (COLUMN_NAMES.size() <= COLUMN_NAME_LIMIT) {
		    COLUMN_NAMES.add(entry.toString());
		    entry.deleteCharAt(1);
		} else {
		    break;
		}
	    }
	    if (COLUMN_NAMES.size() >= COLUMN_NAME_LIMIT) {
		break;
	    }
	    entry = new StringBuilder();
	}
    }

    public static CellStyle getStyle(Workbook wb, StyleType type) {
	Map<StyleType, CellStyle> wbStyles = STYLES.get(wb);

	if (wbStyles == null) {
	    wbStyles = new HashMap<StyleType, CellStyle>();
	    STYLES.put(wb, wbStyles);
	}

	CellStyle style = wbStyles.get(type);

	if (style == null) {
	    boolean isXSSF = wb instanceof XSSFWorkbook || wb instanceof SXSSFWorkbook;
	    Map<FontType, Font> wbFonts = FONTS.get(wb);

	    if (wbFonts == null) {
		wbFonts = new HashMap<FontType, Font>();
		FONTS.put(wb, wbFonts);

		Font bold = wb.createFont();
		bold.setFontName("Arial");
		bold.setBold(true);
		bold.setFontHeight((short) (20 * 10)); // font size=10
		wbFonts.put(FontType.bold, bold);

		Font boldCalibri = wb.createFont();
		boldCalibri.setFontName("Calibri");
		boldCalibri.setBold(true);
		boldCalibri.setFontHeight((short) (20 * 11)); // font size=11
		wbFonts.put(FontType.boldCalibri, boldCalibri);

		Font calibri = wb.createFont();
		calibri.setFontName("Calibri");
		calibri.setFontHeight((short) (20 * 11)); // font size=11
		wbFonts.put(FontType.calibri, calibri);

		Font orangeBold = wb.createFont();
		orangeBold.setFontName("Arial");
		orangeBold.setBold(true);
		if (isXSSF) {
		    ((XSSFFont) orangeBold).setColor(ORANGE);
		} else {
		    orangeBold.setColor(IndexedColors.ORANGE.index);
		}
		orangeBold.setFontHeight((short) (20 * 10)); // font size=10
		wbFonts.put(FontType.orangeBold, orangeBold);

		Font whiteBold = wb.createFont();
		whiteBold.setFontName("Arial");
		whiteBold.setBold(true);
		whiteBold.setColor(IndexedColors.WHITE.index);
		whiteBold.setFontHeight((short) (20 * 10)); // font size=10
		wbFonts.put(FontType.whiteBold, whiteBold);

		Font whiteBoldCalibri = wb.createFont();
		whiteBoldCalibri.setFontName("Calibri");
		whiteBoldCalibri.setBold(true);
		whiteBoldCalibri.setColor(IndexedColors.WHITE.index);
		whiteBoldCalibri.setFontHeight((short) (20 * 11)); // font
								   // size=11
		wbFonts.put(FontType.whiteBoldCalibri, whiteBoldCalibri);

		Font blackBold = wb.createFont();
		blackBold.setFontName("Arial");
		blackBold.setBold(true);
		blackBold.setColor(IndexedColors.BLACK.index);
		blackBold.setFontHeight((short) (20 * 10)); // font size=10
		wbFonts.put(FontType.blackBold, blackBold);

		Font black = wb.createFont();
		black.setFontName("Arial");
		black.setBold(true);
		black.setColor(IndexedColors.BLACK.index);
		black.setFontHeight((short) (20 * 10)); // font size=10
		wbFonts.put(FontType.black, black);

		Font blue = wb.createFont();
		blue.setFontName("Arial");
		blue.setBold(true);
		if (isXSSF) {
		    ((XSSFFont) blue).setColor(BLUE);
		} else {
		    blue.setColor(IndexedColors.BLUE.index);
		}
		blue.setFontHeight((short) (20 * 10)); // font size=10
		wbFonts.put(FontType.blue, blue);

		Font red = wb.createFont();
		red.setFontName("Arial");
		red.setBold(true);
		red.setColor(IndexedColors.RED.index);
		wbFonts.put(FontType.red, red);

		Font redCalibri = wb.createFont();
		redCalibri.setFontName("Calibri");
		redCalibri.setBold(true);
		if (isXSSF) {
		    ((XSSFFont) redCalibri).setColor(RED);
		} else {
		    redCalibri.setColor(IndexedColors.RED.index);
		}
		wbFonts.put(FontType.redCalibri, redCalibri);

		Font redCalibriBold = wb.createFont();
		redCalibriBold.setFontName("Calibri");
		redCalibriBold.setBold(true);
		if (isXSSF) {
		    ((XSSFFont) redCalibriBold).setColor(RED);
		} else {
		    redCalibriBold.setColor(IndexedColors.RED.index);
		}
		wbFonts.put(FontType.redCalibriBold, redCalibriBold);

		Font green = wb.createFont();
		green.setFontName("Arial");
		green.setBold(true);
		green.setColor(IndexedColors.GREEN.index);
		wbFonts.put(FontType.green, green);
	    }

	    Font bold = wbFonts.get(FontType.bold);
	    Font boldCalibri = wbFonts.get(FontType.boldCalibri);
	    Font calibri = wbFonts.get(FontType.calibri);
	    Font orangeBold = wbFonts.get(FontType.orangeBold);
	    Font whiteBold = wbFonts.get(FontType.whiteBold);
	    Font whiteBoldCalibri = wbFonts.get(FontType.whiteBoldCalibri);
	    Font blackBold = wbFonts.get(FontType.blackBold);
	    Font black = wbFonts.get(FontType.black);
	    Font blue = wbFonts.get(FontType.blue);
	    Font red = wbFonts.get(FontType.red);
	    Font redCalibri = wbFonts.get(FontType.redCalibri);
	    Font redCalibriBold = wbFonts.get(FontType.redCalibriBold);
	    Font green = wbFonts.get(FontType.green);

	    DataFormat df = wb.createDataFormat();

	    style = wb.createCellStyle();
	    switch (type) {
	    case styleDate:
		style.setDataFormat(df.getFormat("dd/MM/yyyy"));
		break;
	    case styleMonth:
		style.setDataFormat(df.getFormat("MM/yy"));
		break;
	    case styleTime:
		style.setDataFormat(df.getFormat("HH:mm"));
		break;
	    case styleRed:
		style.setFillForegroundColor(IndexedColors.RED.index);
		setThinBorder(style);
		break;
	    case styleBoldBorder:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setFont(boldCalibri);
		break;
	    case styleBoldBorderBottom:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setFont(boldCalibri);
		break;
	    case styleYellow:
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		setThinBorder(style);
		break;
	    case styleNormalArial:
		style.setFont(black);
		break;
	    case styleYellowLeftBorderRightNoBorder:
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.THIN);
		break;
	    case styleLeftBorderRightNoBorder:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.THIN);
		break;
	    case styleRightBorderLeftNoBorder:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.THIN);
		break;
	    case styleYellowRightBorderLeftNoBorder:
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderTop(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.THIN);
		break;
	    case styleYellowRightBorder:
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		break;
	    case styleYellowLeftBorder:
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		break;
	    case styleYellowLeftRightBorder:
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		break;
	    case styleBold:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setFont(bold);
		break;
	    case styleBoldCenter:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setFont(bold);
		style.setAlignment(HorizontalAlignment.CENTER);
		break;
	    case styleBoldTopBottom:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setFont(bold);
		break;
	    case styleBoldTopLeftBottomBorder:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setFont(bold);
		break;
	    case styleBoldTopRightBottomBorder:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setTopBorderColor(IndexedColors.BLACK.index);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setFont(bold);
		break;
	    case styleBoldPercent:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setFont(bold);
		style.setDataFormat(df.getFormat("#0.00%"));
		break;
	    case stylePercent:
		style.setDataFormat(df.getFormat("#0.00%"));
		break;
	    case stylePercentNoDetails:
		style.setDataFormat(df.getFormat("#0%"));
		break;
	    case stylePercentNoDetailsRightBorder:
		style.setDataFormat(df.getFormat("#0%"));
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		break;
	    case stylePercentBorder:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setDataFormat(df.getFormat("#0.00%"));
		setThinBorder(style);
		break;
	    case stylePercentRightMediumBorderNoDetails:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setDataFormat(df.getFormat("#0%"));
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		break;
	    case stylePercentBorderNoDetails:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setDataFormat(df.getFormat("#0%"));
		setThinBorder(style);
		break;
	    case stylePercentBorderBottomNoDetails:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setDataFormat(df.getFormat("#0%"));
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		break;
	    case stylePercentDetailed:
		style.setDataFormat(df.getFormat("#0.0000%"));
		break;
	    case styleNumber:
		style.setDataFormat(df.getFormat("#0.00"));
		break;
	    case styleNumberBorder:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setDataFormat(df.getFormat("#0.00"));
		break;
	    case styleNumberDezimal1Border:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setDataFormat(df.getFormat("#0.0"));
		break;
	    case styleNumberDezimal3Border:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setDataFormat(df.getFormat("#0.000"));
		break;
	    case styleNumberDezimal4Border:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setDataFormat(df.getFormat("#0.0000"));
		break;
	    case styleNumberBorderGrey:
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
		setThinBorder(style);
		style.setDataFormat(df.getFormat("#0.00"));
		break;
	    case styleNumberBorderRightMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setDataFormat(df.getFormat("#0.00"));
		break;
	    case styleNumberBorderBottomMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setDataFormat(df.getFormat("#0.00"));
		break;
	    case styleNumberBorderBottomRightMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setDataFormat(df.getFormat("#0.00"));
		break;
	    case styleNumberBorderNoDetails:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setDataFormat(df.getFormat("#,##0"));
		break;
	    case styleNumberBorderNoDetailsRed:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setDataFormat(df.getFormat("#,##0"));
		style.setFont(redCalibri);
		break;
	    case styleNumberBoldBorderRightGreyBGNoDetails:
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setFont(boldCalibri);
		style.setDataFormat(df.getFormat("#,##0"));
		break;
	    case styleNumberBoldBorderRightBottomGreyBGNoDetails:
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setFont(boldCalibri);
		style.setDataFormat(df.getFormat("#,##0"));
		break;
	    case styleNumberBorderBottomMediumNoDetails:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setDataFormat(df.getFormat("#,##0"));
		break;
	    case styleNumberBorderBottomMediumNoDetailsRed:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setDataFormat(df.getFormat("#,##0"));
		style.setFont(redCalibri);
		break;
	    case styleNumberBorderBottomRightMediumNoDetails:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setDataFormat(df.getFormat("#,##0"));
		break;
	    case styleNumberBorderBottomRightMediumNoDetailsRed:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setDataFormat(df.getFormat("#,##0"));
		style.setFont(redCalibri);
		break;
	    case styleNumberBorderRightMediumNoDetails:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setDataFormat(df.getFormat("#,##0"));
		break;
	    case styleNumberBorderRightMediumNoDetailsRed:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setDataFormat(df.getFormat("#,##0"));
		style.setFont(redCalibri);
		break;
	    case styleNumberBorderMediumLeftNoDetails:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setDataFormat(df.getFormat("#,##0"));
		break;
	    case styleNumberBorderMediumRightNoDetails:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setDataFormat(df.getFormat("#,##0"));
		break;
	    case styleNumberDetailedBorder:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setDataFormat(df.getFormat("#,##0.00"));
		break;
	    case styleNumberDetailed:
		style.setFont(black);
		style.setDataFormat(df.getFormat("#,##0.00"));
		break;
	    case styleNumberDetailedBold:
		style.setDataFormat(df.getFormat("#,##0.00"));
		style.setFont(bold);
		break;
	    case titleGrey:
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setFont(bold);
		break;
	    case titleGreyAllMedium:
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setFont(bold);
		break;
	    case styleLeftBorder:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		break;
	    case styleLeftBorderGreen:
		style.setFillForegroundColor(IndexedColors.GREEN.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		break;
	    case styleLeftBorderGrey:
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		break;
	    case styleLeftBorderOrange:
		style.setFillForegroundColor(IndexedColors.ORANGE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		break;
	    case styleLeftBorderRed:
		style.setFillForegroundColor(IndexedColors.RED.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		break;
	    case styleLeftRightBorder:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		break;
	    case styleBoldNoBorder:
		style.setFont(bold);
		break;
	    case styleNoBorder:

		break;
	    case styleNoBorderBold:
		style.setFont(boldCalibri);
		break;
	    case styleYellowBlackBoldCenter:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillBackgroundColor(IndexedColors.BLACK.index);
		setThinBorder(style);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(orangeBold);
		break;
	    case styleYellowBlackBoldRight:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillBackgroundColor(IndexedColors.BLACK.index);
		setThinBorder(style);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(orangeBold);
		break;
	    case styleOrangeHeader:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderTop(BorderStyle.THIN);
		if (isXSSF) {
		    ((XSSFCellStyle) style).setTopBorderColor(ORANGE);
		} else {
		    style.setTopBorderColor(IndexedColors.ORANGE.index);
		}
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(orangeBold);
		break;
	    case styleWhiteBlackBoldHeader:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(whiteBold);
		break;
	    case styleWhiteBlackBoldHeaderCenter:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(whiteBold);
		break;
	    case styleWhiteBlackBoldHeaderCenterCalibri:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setFont(whiteBoldCalibri);
		break;
	    case styleWhiteBlackBoldHeaderLeftCalibri:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(whiteBoldCalibri);
		break;
	    case styleWhiteBlackBoldHeaderLeftCalibriTopText:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(whiteBoldCalibri);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		break;
	    case styleWhiteBlackBoldHeaderLeftCalibriWrap:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(whiteBoldCalibri);
		style.setWrapText(true);
		style.setVerticalAlignment(VerticalAlignment.TOP);
		break;
	    case styleWhiteBlackBoldHeaderRightCalibri:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(whiteBoldCalibri);
		break;
	    case styleWhiteBlackBoldNumberRight:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setDataFormat(df.getFormat("#,##0"));
		style.setFont(whiteBold);
		break;
	    case styleWhiteBlackBoldPercentRight:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setDataFormat(df.getFormat("#0%"));
		style.setFont(whiteBold);
		break;
	    case styleBoldTopRightMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blackBold);
		break;
	    case styleRightMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(black);
		break;
	    case styleBoldBottomRightMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blackBold);
		break;
	    case styleBottomRightMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(black);
		break;
	    case styleBottomMediumCalibri:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(calibri);
		break;
	    case blue:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(LIGHT_BLUE);
		} else {
		    style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.WHITE.index);
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.WHITE.index);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blue);
		break;
	    case blueNumber:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(LIGHT_BLUE);
		} else {
		    style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.WHITE.index);
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.WHITE.index);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blue);
		style.setDataFormat(df.getFormat("#0.00"));
		break;
	    case white:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderLeft(BorderStyle.THIN);
		if (isXSSF) {
		    ((XSSFCellStyle) style).setLeftBorderColor(LIGHT_BLUE);
		    ((XSSFCellStyle) style).setRightBorderColor(LIGHT_BLUE);
		} else {
		    style.setLeftBorderColor(IndexedColors.LIGHT_BLUE.index);
		    style.setRightBorderColor(IndexedColors.LIGHT_BLUE.index);
		}
		style.setBorderRight(BorderStyle.THIN);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blue);
		break;
	    case whiteNumber:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderLeft(BorderStyle.THIN);
		if (isXSSF) {
		    ((XSSFCellStyle) style).setLeftBorderColor(LIGHT_BLUE);
		    ((XSSFCellStyle) style).setRightBorderColor(LIGHT_BLUE);
		} else {
		    style.setLeftBorderColor(IndexedColors.LIGHT_BLUE.index);
		    style.setRightBorderColor(IndexedColors.LIGHT_BLUE.index);
		}
		style.setBorderRight(BorderStyle.THIN);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blue);
		style.setDataFormat(df.getFormat("#0.00"));
		break;
	    case styleLeftMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(black);
		break;
	    case styleAllThin:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(black);
		break;
	    case styleBoldBottomLeftMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blackBold);
		break;
	    case styleBoldAllMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blackBold);
		break;
	    case styleAllMediumCalibri:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(calibri);
		break;
	    case styleBottomLeftMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(black);
		break;
	    case styleBoldBottomMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blackBold);
		break;
	    case styleBottomMedium:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(black);
		break;
	    case styleBoldAllThin:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blackBold);
		break;
	    case styleBoldLeft:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blackBold);
		break;
	    case styleBoldRight:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.LEFT);
		style.setFont(blackBold);
		break;
	    case styleOrangeHeaderRight:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderTop(BorderStyle.THIN);
		if (isXSSF) {
		    ((XSSFCellStyle) style).setTopBorderColor(ORANGE);
		} else {
		    style.setTopBorderColor(IndexedColors.ORANGE.index);
		}
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(orangeBold);
		break;
	    case styleOrangeFooter:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillBackgroundColor(IndexedColors.BLACK.index);
		setThinBorder(style);
		style.setFont(orangeBold);
		break;
	    case styleBoldLeftBoldBorder:
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		style.setFont(bold);
		break;
	    case styleGreyBoldLeftRightBoldBorder:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setFont(bold);
		break;
	    case styleGreyBoldLeftRightBoldBorderCalibri:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setFont(boldCalibri);
		break;
	    case styleGreyBoldRightLeftRightBoldBorderCalibri:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setFont(boldCalibri);
		style.setAlignment(HorizontalAlignment.RIGHT);
		break;
	    case styleGreyBoldRightBoldBorderCalibri:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setFont(boldCalibri);
		break;
	    case styleGreyBoldLeftCalibri:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		break;
	    case styleGreyBoldRightCalibri:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		style.setAlignment(HorizontalAlignment.RIGHT);
		break;
	    case styleGreyBoldRightCalibriBottom:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleGreyBoldRightCalibriNumber:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setDataFormat(df.getFormat("#0.00"));
		break;
	    case styleGreyBoldRightCalibriNumberBottom:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setDataFormat(df.getFormat("#0.00"));
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleGreyBoldLeftCalibriBottom:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleGreyRightBorderNoDetailsBold:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		style.setDataFormat(df.getFormat("#0%"));
		break;
	    case styleGreyRightBottomBorderNoDetailsBold:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setDataFormat(df.getFormat("#0%"));
		break;
	    case styleBoldLeftBoldBorderYellowBg:
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		style.setFont(bold);
		break;
	    case styleBoldLeftBoldBorderRedBg:
		style.setFillForegroundColor(IndexedColors.ROSE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		style.setFont(bold);
		break;
	    case styleBoldLeftBottomBoldBorder:
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		style.setFont(bold);
		break;
	    case styleLeftBottomBoldBorder:
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleBottomBoldBorder:
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleRightBoldBorderBottomThin:
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleRightBoldBorderBottomThinYellowBg:
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleRightBoldBorderBottomThinRedBg:
		style.setFillForegroundColor(IndexedColors.ROSE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleBottomBorder:
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleTopBottomBorder:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderTop(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.index);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleBoldBottomRight:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setBorderTop(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleBoldBottomRightRed:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setFont(red);
		break;
	    case styleBoldBottom:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		break;
	    case styleBoldGreyLeftBold:
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		style.setTopBorderColor(IndexedColors.BLACK.index);
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setFont(bold);
		break;
	    case styleBoldGrey:
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setTopBorderColor(IndexedColors.BLACK.index);
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setFont(bold);
		break;
	    case styleBoldGreyRightBold:
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setTopBorderColor(IndexedColors.BLACK.index);
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setFont(bold);
		break;
	    case styleLeftBoldBorder:
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleRightBoldBorder:
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleLeftTopRightBoldBorder:
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleLeftTopBoldBorder:
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleTopRightBoldBorder:
		style.setBorderTop(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleLeftBottomRightBoldBorder:
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleLeftBottomBoldBorder2:
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleBottomRightBoldBorder:
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleLeftRightBoldBorder:
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleLeftRightBoldBorderRight:
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setAlignment(HorizontalAlignment.RIGHT);
		break;
	    case styleLeftRightBoldBorderRightGrey:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setAlignment(HorizontalAlignment.RIGHT);
		break;
	    case styleLeftRightBoldBorderRightGreyPercent:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setDataFormat(df.getFormat("#0%"));
		break;
	    case styleLeftRightBoldBorderRightBottomGreyPercent:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setDataFormat(df.getFormat("#0%"));
		break;
	    case styleLeftBoldBorderRightBottomGrey:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFont(boldCalibri);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setAlignment(HorizontalAlignment.RIGHT);
		break;
	    case styleLeftRightBoldBorderRightRedCalibri:
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(redCalibri);
		break;
	    case styleLeftRightBoldBorderRightRedCalibriGrey:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(redCalibriBold);
		break;
	    case styleLeftBottomBoldBorderRightRedCalibriGrey:
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(GREY_50_PERCENT);
		} else {
		    style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.index);
		}
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setAlignment(HorizontalAlignment.RIGHT);
		style.setFont(redCalibriBold);
		break;
	    case styleRightBoldBorderRed:
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setFont(red);
		break;
	    case styleRightBoldBorderBottomThinRed:
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setFont(red);
		break;
	    case styleOrangeFooterRed:
		style.setFillForegroundColor(IndexedColors.BLACK.index);
		style.setFillBackgroundColor(IndexedColors.BLACK.index);
		setThinBorder(style);
		style.setFont(red);
		break;
	    case styleRedBg:
		style.setFillForegroundColor(IndexedColors.ROSE.index);
		setThinBorder(style);
		break;
	    case styleYellowBg:
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		setThinBorder(style);
		break;
	    case styleYellowTopBottomBorder:
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleDate2:
		style.setDataFormat(df.getFormat("dd/MM/yyyy"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		break;
	    case styleTime2:
		style.setDataFormat(df.getFormat("HH:mm:ss"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		break;
	    case styleTime2NoBorder:
		style.setDataFormat(df.getFormat("HH:mm:ss"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		break;
	    case styleTimeBorder:
		style.setDataFormat(df.getFormat("HH:mm"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		break;
	    case styleTimeBorderGreen:
		style.setDataFormat(df.getFormat("HH:mm"));
		style.setFillForegroundColor(IndexedColors.GREEN.index);
		setThinBorder(style);
		break;
	    case styleDateTime:
		style.setDataFormat(df.getFormat("dd/MM/yyyy HH:mm"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		break;
	    case styleDateTimeWithSeconds:
		style.setDataFormat(df.getFormat("dd/MM/yyyy HH:mm:ss"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		break;
	    case styleDateTimeWithSecondsBorder:
		style.setDataFormat(df.getFormat("dd/MM/yyyy HH:mm:ss"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		break;
	    case styleTimeWithSecondsBorder:
    	style.setDataFormat(df.getFormat("HH:mm:ss"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		break;
	    case styleDateTimeArial:
		style.setDataFormat(df.getFormat("dd/MM/yyyy HH:mm"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFont(black);
		break;
	    case styleDateTimeBorder:
		style.setDataFormat(df.getFormat("dd/MM/yyyy HH:mm"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		break;
	    case styleDateTimeBorderGreen:
		style.setDataFormat(df.getFormat("dd/MM/yyyy HH:mm"));
		style.setFillForegroundColor(IndexedColors.GREEN.index);
		setThinBorder(style);
		break;
	    case styleDateBorder:
		style.setDataFormat(df.getFormat("dd/MM/yyyy"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		break;
	    case styleDateBorderGreen:
		style.setDataFormat(df.getFormat("dd/MM/yyyy"));
		style.setFillForegroundColor(IndexedColors.GREEN.index);
		setThinBorder(style);
		case styleDateBorderLightYellow:
		style.setDataFormat(df.getFormat("dd/MM/yyyy"));
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
		setThinBorder(style);
		break;
		case styleTimeBorderLightYellow:
			style.setDataFormat(df.getFormat("HH:mm"));
			style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
			setThinBorder(style);
			break;
	    case styleDateTimeBorderBottomMedium:
		style.setDataFormat(df.getFormat("dd/MM/yyyy HH:mm"));
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		style.setBorderBottom(BorderStyle.MEDIUM);
		break;
	    case styleDateTimeOrange:
		style.setDataFormat(df.getFormat("dd/MM/yyyy HH:mm"));
		style.setFillForegroundColor(IndexedColors.ORANGE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(ORANGE);
		} else {
		    style.setFillForegroundColor(IndexedColors.ORANGE.index);
		}
		break;
	    case styleDateTimeRed:
		style.setDataFormat(df.getFormat("dd/MM/yyyy HH:mm"));
		style.setFillForegroundColor(IndexedColors.RED.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(RED);
		} else {
		    style.setFillForegroundColor(IndexedColors.RED.index);
		}
		break;
	    case styleRedArial:
		style.setFont(black);
		style.setFillForegroundColor(IndexedColors.RED.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(RED);
		} else {
		    style.setFillForegroundColor(IndexedColors.RED.index);
		}
		break;
	    case styleTopLeftBottom:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		style.setBorderTop(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleTopBottom:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderTop(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleTopRightBottom:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBottomBorderColor(IndexedColors.BLACK.index);
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.BLACK.index);
		style.setBorderTop(BorderStyle.THIN);
		style.setTopBorderColor(IndexedColors.BLACK.index);
		break;
	    case styleRightLeft:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setBorderLeft(BorderStyle.THIN);
		style.setLeftBorderColor(IndexedColors.BLACK.index);
		style.setBorderRight(BorderStyle.THIN);
		style.setRightBorderColor(IndexedColors.BLACK.index);
	    case styleBoldLeftRight:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setThinBorder(style);
		style.setBorderLeft(BorderStyle.MEDIUM);
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setAlignment(HorizontalAlignment.RIGHT);
		break;
	    case styleArial:
		style.setFont(black);
		break;
	    case styleBoldCenterLightGreen:
    	style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		setMediumBorder(style);
		style.setFont(bold);
		style.setAlignment(HorizontalAlignment.CENTER);
    	break;
    	
	    case styleBoldCenterLightYellow:
	    	style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			setMediumBorder(style);
			style.setFont(bold);
			style.setAlignment(HorizontalAlignment.CENTER);
	    	break;
	    case styleBoldCenterLightRed:
	    	style.setFillForegroundColor(IndexedColors.RED.index);
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			setMediumBorder(style);
			style.setFont(bold);
			style.setAlignment(HorizontalAlignment.CENTER);
	    	break;
    	
    	
	    case styleArialWrapText:
	    	style.setFont(black);
	    	style.setWrapText(true);
	    	break;

	    case styleBottomRightBorder:
			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			style.setBottomBorderColor(IndexedColors.BLACK.index);
			style.setRightBorderColor(IndexedColors.BLACK.index);
			break;
		
	    case styleRightBorder:
			style.setBorderRight(BorderStyle.THIN);
			style.setRightBorderColor(IndexedColors.BLACK.index);
	    	break;
	    	
	    case styleLeftBorderRightBorder:
			style.setFillForegroundColor(IndexedColors.WHITE.index);
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			break;

	    case styleLeftBorderRightBorderBottomBorder:
			style.setFillForegroundColor(IndexedColors.WHITE.index);
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			style.setBorderBottom(BorderStyle.THIN);
			break;
	    case stylePercentDetailed1DecimalPlace:
		style.setDataFormat(df.getFormat("#0.0%"));
		break;
        
	    case stylePercentBorderGrey:
            style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
            style.setDataFormat(df.getFormat("#0.00%"));
            setThinBorder(style);
            break;
	    case styleGreenNoBorder:
	    	style.setFont(green);
	    	break;
	    case styleRedNoBorder:
	    	style.setFont(red);
	    	break;
	    case styleOrangeBorder:
		style.setFillForegroundColor(IndexedColors.ORANGE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		if (isXSSF) {
		    ((XSSFCellStyle) style).setFillForegroundColor(ORANGE);
		} else {
		    style.setFillForegroundColor(IndexedColors.ORANGE.index);
		}
        setThinBorder(style);
		break;
	    case styleTimeOrangeBorder:
			style.setDataFormat(df.getFormat("HH:mm"));
			style.setFillForegroundColor(IndexedColors.ORANGE.index);
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			if (isXSSF) {
			    ((XSSFCellStyle) style).setFillForegroundColor(ORANGE);
			} else {
			    style.setFillForegroundColor(IndexedColors.ORANGE.index);
			}
            setThinBorder(style);
			break;
	    case styleDateOrangeBorder:
    		style.setDataFormat(df.getFormat("dd/MM/yyyy"));
    		style.setFillForegroundColor(IndexedColors.ORANGE.index);
    		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    		if (isXSSF) {
    		    ((XSSFCellStyle) style).setFillForegroundColor(ORANGE);
    		} else {
    		    style.setFillForegroundColor(IndexedColors.ORANGE.index);
    		}
            setThinBorder(style);
    		break;
	    case styleDateGermanStandardCentered:
			style.setDataFormat(df.getFormat("dd.MM.yyyy"));
			style.setAlignment(HorizontalAlignment.CENTER);
	    	break;
	    case styleFillColorLightGreenBoldCentered:
	    	style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
	    	style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			style.setFont(bold);
	    	style.setAlignment(HorizontalAlignment.CENTER);
	    	break;
	    case styleNumberLightGreenBoldCentered:
	    	style.setDataFormat(df.getFormat("#0.00"));
	    	style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
	    	style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    	style.setFont(bold);
	    	style.setAlignment(HorizontalAlignment.CENTER);
			break;
	    case styleNumberRedBoldCentered:
	    	style.setDataFormat(df.getFormat("#0.00"));
	    	style.setFillForegroundColor(IndexedColors.RED.index);
	    	style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    	style.setAlignment(HorizontalAlignment.CENTER);
			style.setFont(bold);
			break;
	    case styleRoseNoBorderBoldCentered:
	    	style.setFillForegroundColor(IndexedColors.ROSE.index);
	    	style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	    	style.setFont(bold);
	    	style.setAlignment(HorizontalAlignment.CENTER);
	    	break;
	    case styleCentered:
			style.setAlignment(HorizontalAlignment.CENTER);
			break;
	    case styleRedArialBoldCentered:
			style.setFont(black);
			style.setFillForegroundColor(IndexedColors.RED.index);
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			if (isXSSF) {
			    ((XSSFCellStyle) style).setFillForegroundColor(RED);
			} else {
			    style.setFillForegroundColor(IndexedColors.RED.index);
			}
			style.setFont(bold);
			style.setAlignment(HorizontalAlignment.CENTER);
			break;
	    case styleRightAligned:
	    	style.setAlignment(HorizontalAlignment.RIGHT);
			break;
	    default:
		style.setFillForegroundColor(IndexedColors.WHITE.index);
		setThinBorder(style);
		break;
	    }

	    wbStyles.put(type, style);
	}

	return style;
    }
    
    private static void setThinBorder(CellStyle style) {
	style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	style.setBorderBottom(BorderStyle.THIN);
	style.setBottomBorderColor(IndexedColors.BLACK.index);
	style.setBorderLeft(BorderStyle.THIN);
	style.setLeftBorderColor(IndexedColors.BLACK.index);
	style.setBorderRight(BorderStyle.THIN);
	style.setRightBorderColor(IndexedColors.BLACK.index);
	style.setBorderTop(BorderStyle.THIN);
	style.setTopBorderColor(IndexedColors.BLACK.index);
    }

    private static void setMediumBorder(CellStyle style) {
	style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	style.setBorderBottom(BorderStyle.NONE);
	style.setBottomBorderColor(IndexedColors.BLACK.index);
	style.setBorderLeft(BorderStyle.NONE);
	style.setLeftBorderColor(IndexedColors.BLACK.index);
	style.setBorderRight(BorderStyle.NONE);
	style.setRightBorderColor(IndexedColors.BLACK.index);
	style.setBorderTop(BorderStyle.NONE);
	style.setTopBorderColor(IndexedColors.BLACK.index);
    }

    // rb-1197 added trackAllColumnsForAutoSizing due to library update
    public static void autosizeColumns(Sheet sheet, int startColumn, int endColumn) {
    	autosizeColumns(sheet, startColumn, endColumn, false);
    }

    // RB-14909 provide new method that provides the possibility to consider merged cells
    public static void autosizeColumns(Sheet sheet, int startColumn, int endColumn,
    		boolean considerMergedCellsForAutosizing)
    {
		if (!sheet.getClass().getName().equals("org.apache.poi.hssf.usermodel.HSSFSheet")
				&& !sheet.getClass().getName().equals("org.apache.poi.xssf.usermodel.XSSFSheet")) {
    		SXSSFSheet.class.cast(sheet).trackAllColumnsForAutoSizing();
    	}
		
		for (int i = startColumn; i < endColumn; i++)
		{
			try
			{
				sheet.autoSizeColumn(i, considerMergedCellsForAutosizing);
			}
			catch(Exception e)
			{
				String name="";
				String sheetname="";
				
				if(sheet!=null)
				{
					sheetname=sheet.getSheetName();
					
					Row row=sheet.getRow(2);
					if(row!=null)
					{
						Cell columnname=row.getCell(i);
						if(columnname!=null)
						{
							name=columnname.getStringCellValue();
						}
					}
				}
				LOG.error("Error at autoSizeColum:"+i+ "Sheetname:"+sheetname+" name:"+name,e);
			}
		}
    }

    // with this method you can view all available cell styles in your workbook
    public static void debugAllCellStyles(Workbook wb) {
	Sheet sheet = wb.createSheet("Cell styles");
	CreationHelper createHelper = wb.getCreationHelper();

	int row = 1;
	Row r = sheet.createRow(row);
	Cell c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("style"));
	c.setCellStyle(getStyle(wb, StyleType.style));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleArial"));
	c.setCellStyle(getStyle(wb, StyleType.styleArial));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldBorderBottom"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldBorderBottom));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleYellow"));
	c.setCellStyle(getStyle(wb, StyleType.styleYellow));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleYellowLeftBorderRightNoBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleYellowLeftBorderRightNoBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleYellowRightBorderLeftNoBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleYellowRightBorderLeftNoBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftBorderRightNoBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftBorderRightNoBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRightBorderLeftNoBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleRightBorderLeftNoBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleYellowRightBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleYellowRightBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRed"));
	c.setCellStyle(getStyle(wb, StyleType.styleRed));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBold"));
	c.setCellStyle(getStyle(wb, StyleType.styleBold));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldCenter"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldCenter));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldTopBottom"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldTopBottom));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldTopLeftBottomBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldTopLeftBottomBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldTopRightBottomBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldTopRightBottomBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("titleGrey"));
	c.setCellStyle(getStyle(wb, StyleType.titleGrey));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("titleGreyAllMedium"));
	c.setCellStyle(getStyle(wb, StyleType.titleGreyAllMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleYellowLeftBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleYellowLeftBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleYellowLeftRightBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleYellowLeftRightBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftRightBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftRightBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleBoldPercent));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleBoldPercent"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.stylePercent));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("stylePercent"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.stylePercentNoDetails));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("stylePercentNoDetails"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.stylePercentNoDetailsRightBorder));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("stylePercentNoDetailsRightBorder"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.stylePercentBorder));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("stylePercentBorder"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.stylePercentRightMediumBorderNoDetails));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("stylePercentRightMediumBorderNoDetails"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.stylePercentBorderNoDetails));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("stylePercentBorderNoDetails"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.stylePercentBorderBottomNoDetails));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("stylePercentBorderBottomNoDetails"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleGreyRightBorderNoDetailsBold));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleGreyRightBorderNoDetailsBold"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleGreyRightBottomBorderNoDetailsBold));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleGreyRightBottomBorderNoDetailsBold"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleLeftRightBoldBorderRightGreyPercent));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleLeftRightBoldBorderRightGreyPercent"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleLeftRightBoldBorderRightBottomGreyPercent));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleLeftRightBoldBorderRightBottomGreyPercent"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548754);
	c.setCellStyle(getStyle(wb, StyleType.stylePercentDetailed));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("stylePercentDetailed"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548754);
	c.setCellStyle(getStyle(wb, StyleType.stylePercentDetailed1DecimalPlace));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("stylePercentDetailed1DecimalPlace"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumber));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumber"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorder));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorder"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderRightMedium));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderRightMedium"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberDetailed));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberDetailed"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberDetailedBold));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberDetailedBold"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldNoBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldNoBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleYellowBlackBoldCenter"));
	c.setCellStyle(getStyle(wb, StyleType.styleYellowBlackBoldCenter));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleYellowBlackBoldRight"));
	c.setCellStyle(getStyle(wb, StyleType.styleYellowBlackBoldRight));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleOrangeHeader"));
	c.setCellStyle(getStyle(wb, StyleType.styleOrangeHeader));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleOrangeHeaderRight"));
	c.setCellStyle(getStyle(wb, StyleType.styleOrangeHeaderRight));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldLeftBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldLeftBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleGreyBoldLeftRightBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleGreyBoldLeftRightBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBottomBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleBottomBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRightBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleRightBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBottomRightBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleBottomRightBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleTopBottomBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleTopBottomBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRightBoldBorderBottomThin"));
	c.setCellStyle(getStyle(wb, StyleType.styleRightBoldBorderBottomThin));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleOrangeFooter"));
	c.setCellStyle(getStyle(wb, StyleType.styleOrangeFooter));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldLeftBottomBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldLeftBottomBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBottomBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleBottomBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldBottom"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldBottom));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldBottomRight"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldBottomRight));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldGreyLeftBold"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldGreyLeftBold));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldGrey"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldGrey));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldGreyRightBold"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldGreyRightBold));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRightBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleRightBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRightBoldBorderBottomThinRed"));
	c.setCellStyle(getStyle(wb, StyleType.styleRightBoldBorderBottomThinRed));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldBottomRightRed"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldBottomRightRed));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleOrangeFooterRed"));
	c.setCellStyle(getStyle(wb, StyleType.styleOrangeFooterRed));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRightBoldBorderRed"));
	c.setCellStyle(getStyle(wb, StyleType.styleRightBoldBorderRed));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftBottomBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftBottomBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldLeftBoldBorderRedBg"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldLeftBoldBorderRedBg));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRedBg"));
	c.setCellStyle(getStyle(wb, StyleType.styleRedBg));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRightBoldBorderBottomThinRedBg"));
	c.setCellStyle(getStyle(wb, StyleType.styleRightBoldBorderBottomThinRedBg));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldLeftBoldBorderYellowBg"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldLeftBoldBorderYellowBg));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleYellowBg"));
	c.setCellStyle(getStyle(wb, StyleType.styleYellowBg));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleYellowBottomBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleYellowTopBottomBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRightBoldBorderBottomThinYellowBg"));
	c.setCellStyle(getStyle(wb, StyleType.styleRightBoldBorderBottomThinYellowBg));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftTopRightBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftTopRightBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftTopBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftTopBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleTopRightBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleTopRightBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftBottomRightBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftBottomRightBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftBottomBoldBorder2"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftBottomBoldBorder2));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBottomRightBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleBottomRightBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftRightBoldBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftRightBoldBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftRightBoldBorderRight"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftRightBoldBorderRight));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftRightBoldBorderRightGrey"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftRightBoldBorderRightGrey));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftBoldBorderRightBottomGrey"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftBoldBorderRightBottomGrey));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftRightBoldBorderRightRedCalibri"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftRightBoldBorderRightRedCalibri));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftRightBoldBorderRightRedCalibriGrey"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftRightBoldBorderRightRedCalibriGrey));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftBottomBoldBorderRightRedCalibriGrey"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftBottomBoldBorderRightRedCalibriGrey));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleNoBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleNoBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleNoBorderBold"));
	c.setCellStyle(getStyle(wb, StyleType.styleNoBorderBold));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleWhiteBlackBoldHeader"));
	c.setCellStyle(getStyle(wb, StyleType.styleWhiteBlackBoldHeader));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleWhiteBlackBoldHeaderCenter"));
	c.setCellStyle(getStyle(wb, StyleType.styleWhiteBlackBoldHeaderCenter));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleWhiteBlackBoldHeaderCenterCalibri"));
	c.setCellStyle(getStyle(wb, StyleType.styleWhiteBlackBoldHeaderCenterCalibri));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleWhiteBlackBoldHeaderLeftCalibri"));
	c.setCellStyle(getStyle(wb, StyleType.styleWhiteBlackBoldHeaderLeftCalibri));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleWhiteBlackBoldHeaderLeftCalibriTopText"));
	c.setCellStyle(getStyle(wb, StyleType.styleWhiteBlackBoldHeaderLeftCalibriTopText));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleWhiteBlackBold\nHeaderLeftCalibriWrap"));
	c.setCellStyle(getStyle(wb, StyleType.styleWhiteBlackBoldHeaderLeftCalibriWrap));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleWhiteBlackBoldHeaderRightCalibri"));
	c.setCellStyle(getStyle(wb, StyleType.styleWhiteBlackBoldHeaderRightCalibri));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldTopRightMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldTopRightMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRightMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleRightMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldBottomRightMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldBottomRightMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleAllThin"));
	c.setCellStyle(getStyle(wb, StyleType.styleAllThin));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldBottomLeftMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldBottomLeftMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldAllMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldAllMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleAllMediumCalibri"));
	c.setCellStyle(getStyle(wb, StyleType.styleAllMediumCalibri));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldBottomMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldBottomMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldAllThin"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldAllThin));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldLeft"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldLeft));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldRight"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldRight));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBottomLeftMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleBottomLeftMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBottomMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleBottomMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBottomRightMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleBottomRightMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBottomMediumCalibri"));
	c.setCellStyle(getStyle(wb, StyleType.styleBottomMediumCalibri));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderBottomMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderBottomMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderBottomRightMedium"));
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderBottomRightMedium));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("blue"));
	c.setCellStyle(getStyle(wb, StyleType.blue));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("white"));
	c.setCellStyle(getStyle(wb, StyleType.white));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(27.877d);
	c.setCellStyle(getStyle(wb, StyleType.whiteNumber));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(27.877d);
	c.setCellStyle(getStyle(wb, StyleType.blueNumber));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218d);
	c.setCellStyle(getStyle(wb, StyleType.styleDate));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDate"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218d);
	c.setCellStyle(getStyle(wb, StyleType.styleMonth));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleMonth"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleTime));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleTime"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218d);
	c.setCellStyle(getStyle(wb, StyleType.styleDate2));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDate2"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleDateTime));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDateTime"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleDateTimeWithSeconds));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDateTimeWithSeconds"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleDateTimeArial));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDateTimeArial"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleDateTimeBorder));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDateTimeBorder"));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleDateTimeBorderBottomMedium));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDateTimeBorderBottomMedium"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleDateTimeOrange));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDateTimeOrange"));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleDateOrangeBorder));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDateOrangeBorder"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleTimeOrangeBorder));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleTimeOrangeBorder"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleDateBorderGreen));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDateBorderGreen"));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleDateTimeRed));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDateTimeRed"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleTime2));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleTime2"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleTime2NoBorder));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleTime2NoBorder"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleTimeBorder));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleTimeBorder"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberDetailedBorder));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberDetailedBorder"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderNoDetails));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderNoDetails"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderNoDetailsRed));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderNoDetailsRed"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBoldBorderRightGreyBGNoDetails));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBoldBorderRightGreyBGNoDetails"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBoldBorderRightBottomGreyBGNoDetails));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBoldBorderRightBottomGreyBGNoDetails"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderBottomMediumNoDetails));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderBottomMediumNoDetails"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderBottomMediumNoDetailsRed));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderBottomMediumNoDetailsRed"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderBottomRightMediumNoDetails));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderBottomRightMediumNoDetails"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderBottomRightMediumNoDetailsRed));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderBottomRightMediumNoDetailsRed"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderRightMediumNoDetails));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderRightMediumNoDetails"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderRightMediumNoDetailsRed));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderRightMediumNoDetailsRed"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberBorderMediumLeftNoDetails));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberBorderMediumLeftNoDetails"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleWhiteBlackBoldNumberRight));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleWhiteBlackBoldNumberRight"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548754);
	c.setCellStyle(getStyle(wb, StyleType.styleWhiteBlackBoldPercentRight));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleWhiteBlackBoldPercentRight"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleTopLeftBottom"));
	c.setCellStyle(getStyle(wb, StyleType.styleTopLeftBottom));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRedArial"));
	c.setCellStyle(getStyle(wb, StyleType.styleRedArial));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRedArialBoldCentered"));
	c.setCellStyle(getStyle(wb, StyleType.styleRedArialBoldCentered));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleTopBottom"));
	c.setCellStyle(getStyle(wb, StyleType.styleTopBottom));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleTopRightBottom"));
	c.setCellStyle(getStyle(wb, StyleType.styleTopRightBottom));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRightLeft"));
	c.setCellStyle(getStyle(wb, StyleType.styleRightLeft));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleGreyBoldLeftRightBoldBorderCalibri"));
	c.setCellStyle(getStyle(wb, StyleType.styleGreyBoldLeftRightBoldBorderCalibri));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleGreyBoldRightLeftRightBoldBorderCalibri"));
	c.setCellStyle(getStyle(wb, StyleType.styleGreyBoldRightLeftRightBoldBorderCalibri));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleGreyBoldRightBoldBorderCalibri"));
	c.setCellStyle(getStyle(wb, StyleType.styleGreyBoldRightBoldBorderCalibri));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleGreyBoldLeftCalibri"));
	c.setCellStyle(getStyle(wb, StyleType.styleGreyBoldLeftCalibri));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleGreyBoldRightCalibri"));
	c.setCellStyle(getStyle(wb, StyleType.styleGreyBoldRightCalibri));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleGreyBoldRightCalibriBottom"));
	c.setCellStyle(getStyle(wb, StyleType.styleGreyBoldRightCalibriBottom));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleGreyBoldRightCalibriNumber));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleGreyBoldRightCalibriNumber"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(123456.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleGreyBoldRightCalibriNumberBottom));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleGreyBoldRightCalibriNumberBottom"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleGreyBoldLeftCalibriBottom"));
	c.setCellStyle(getStyle(wb, StyleType.styleGreyBoldLeftCalibriBottom));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleNormalArial"));
	c.setCellStyle(getStyle(wb, StyleType.styleNormalArial));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleBoldCenterLightGreen"));
	c.setCellStyle(getStyle(wb, StyleType.styleBoldCenterLightGreen));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleArial\nWrapText"));
	c.setCellStyle(getStyle(wb, StyleType.styleArialWrapText));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftBorderRightBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftBorderRightBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleLeftBorderRightBorderBottomBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleLeftBorderRightBorderBottomBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.stylePercentBorderGrey));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("stylePercentBorderGrey"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleGreenNoBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleGreenNoBorder));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRedNoBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleRedNoBorder));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleDateBorderLightYellow));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDateBorderLightYellow"));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleTimeBorderLightYellow));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleTimeBorderLightYellow"));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleOrangeBorder"));
	c.setCellStyle(getStyle(wb, StyleType.styleOrangeBorder));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(41218.75d);
	c.setCellStyle(getStyle(wb, StyleType.styleDateGermanStandardCentered));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleDateGermanStandard"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleFillColorLightGreenBoldCentered"));
	c.setCellStyle(getStyle(wb, StyleType.styleFillColorLightGreenBoldCentered));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberLightGreenBoldCentered));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberLightGreenBoldCentered"));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(.2548);
	c.setCellStyle(getStyle(wb, StyleType.styleNumberRedBoldCentered));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleNumberRedBoldCentered"));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellValue(createHelper.createRichTextString("styleRoseNoBorderBoldCentered"));
	c.setCellStyle(getStyle(wb, StyleType.styleRoseNoBorderBoldCentered));
	row += 2;

	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellStyle(getStyle(wb, StyleType.styleCentered));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleCentered"));
	row += 2;
	
	r = sheet.createRow(row);
	c = r.createCell(1);
	c.setCellStyle(getStyle(wb, StyleType.styleRightAligned));
	c = r.createCell(2);
	c.setCellValue(createHelper.createRichTextString("styleRightAligned"));
	row += 2;
	
	
	
	autosizeColumns(sheet, 1, 3);
    }

    public static String getIndexChar(int index) {
	String result = COLUMN_NAMES.get(index);
	if (result == null) {
	    LOG.error("index out of bounds: " + index);
	}

	return result;
    }

    public static void setRballColors(HSSFWorkbook wb) {
	// r/g/b values
	HSSFPalette palette = wb.getCustomPalette();

	// rball orange
	palette.setColorAtIndex(IndexedColors.ORANGE.index, (byte) COLOR_ORANGE.getRed(), (byte) COLOR_ORANGE.getGreen(), (byte) COLOR_ORANGE.getBlue());

	// light blue
	palette.setColorAtIndex(IndexedColors.LIGHT_BLUE.index, (byte) COLOR_LIGHT_BLUE.getRed(), (byte) COLOR_LIGHT_BLUE.getGreen(), (byte) COLOR_LIGHT_BLUE.getBlue());

	// blue
	palette.setColorAtIndex(IndexedColors.BLUE.index, (byte) COLOR_BLUE.getRed(), (byte) COLOR_BLUE.getGreen(), (byte) COLOR_BLUE.getBlue());

	// very light grey
	palette.setColorAtIndex(IndexedColors.GREY_50_PERCENT.index, (byte) COLOR_GREY_50_PERCENT.getRed(), (byte) COLOR_GREY_50_PERCENT.getGreen(), (byte) COLOR_GREY_50_PERCENT.getBlue());
    }

    public static double getNumValue(Cell cell) {
	double result = 0d;

	try {
	    result = cell.getNumericCellValue();
	} catch (Exception e) {
	    LOG.warn("error:\n" + e.getMessage());
	    LOG.warn("trying to parse value from String...");
	    String value = cell.getStringCellValue();
	    if (DataFormatter.getInstance().isDoubleNumber(value)) {
		result = Double.parseDouble(value);
	    }
	}

	return result;
    }

    public static void clearWorkbookData(Workbook wb) {
	FONTS.remove(wb);
	STYLES.remove(wb);
    }

    /**
     * @return the orange
     */
    public static XSSFColor getOrange() {
	return ORANGE;
    }

    /**
     * @return the lightBlue
     */
    public static XSSFColor getLightBlue() {
	return LIGHT_BLUE;
    }

    /**
     * @return the blue
     */
    public static XSSFColor getBlue() {
	return BLUE;
    }

    /**
     * @return the grey50Percent
     */
    public static XSSFColor getGrey50Percent() {
	return GREY_50_PERCENT;
    }
}