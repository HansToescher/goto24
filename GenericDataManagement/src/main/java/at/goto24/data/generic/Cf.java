package at.goto24.data.generic;

/**
 * Field names for ItemRecord management...
 * @author j.toescher
 * @version 1.0.0
 */
public class Cf {
	
	// Data
	public static final String DATA_COLUMS = "_columns";
	public static final String DATA_TOTAL = "_total";
	public static final String DATA_GROUP = "_group";
	public static final String BASE_DIR = "baseDir"; // File Object
	
	// Parameter
	public static final String PAR_DATE_TO = "PAR_DATE_TO";
	public static final String PAR_DATE_FROM = "PAR_DATE_FROM";
	public static final String PAR_QUERY = "PAR_QUERY";
	public static final String PAR_SHEET_TITLE = "PAR_SHEET_TITLE";
	public static final String PAR_GROUP_ID = "PAR_GROUP_ID";
	public static final String PAR_REPLACE_PARAMS = "PAR_REPLACE_PARAMS";
	public static final String PAR_ORDER_BY = "PAR_ORDER_BY";
	public static final String PAR_FIELD_FORMAT = "PAR_FIELD_FORMAT";
	public static final String PAR_TOTAL_FIELDS = "PAR_TOTAL_FIELDS";
	public static final String PAR_TOTAL_TEXT = "PAR_TOTAL_TEXT";
	public static final String PAR_NO_EXCEL_FILE = "PAR_NO_EXCEL_FILE";
	public static final String PAR_HIDE_FIELDS = "PAR_HIDE_FIELDS";
	public static final String PAR_TYPE_OF_FIELD = "PAR_TYPE_OF_FIELD";
	public static final String PAR_SQL_TO_ORACLE = "PAR_SQL_TO_ORACLE";
	public static final String PAR_SQL_TO_MYSQL = "PAR_SQL_TO_MYSQL";
	// optional 
	public static final String PAR_EXCEL_FILE_NAME = "PAR_EXCEL_FILE_NAME";
	// Group Operations
	public static final String PAR_GROUP_FIELD = "PAR_GROUP_FIELD";
	public static final String PAR_GROUP_SHEET_PREFIX = "PAR_GROUP_SHEET_PREFIX";
	public static final String PAR_GROUP_SHEET_SUMMARY = "PAR_GROUP_SHEET_SUMMARY";
	public static final String PAR_GROUP_SHOW_DETAIL_SHEET = "PAR_GROUP_SHOW_DETAIL_SHEET";
	public static final String PAR_GROUP_SHEET_ORDER_BY = "PAR_GROUP_SHEET_ORDER_BY";
	public static final String PAR_GROUP_RESULT_ORDER_BY = "PAR_GROUP_RESULT_ORDER_BY";
	public static final String PAR_TRANSLATE_HEADER = "PAR_TRANSLATE_HEADER";
	public static final String PAR_CONVERT_HEADER_CASES = "PAR_CONVERT_HEADER_CASES";
	public static final String PAR_TRANSLATE_FIELD_VALUE = "PAR_TRANSLATE_FIELD_VALUE";
	public static final String PAR_FIELD_TIMEZONE = "PAR_FIELD_TIMEZONE";
	public static final String PAR_GROUP_SHEET_HIDE_FIELDS = "PAR_GROUP_SHEET_HIDE_FIELDS";
	public static final String PAR_FIELD_TO_UPPERCASE = "PAR_FIELD_TO_UPPERCASE";
	// only used by GenericSubQueryFieldsReport.class
	public static final String PAR_FIELD_OF_SUB_QUERIES = "PAR_FIELD_OF_SUB_QUERIES";
	public static final String PAR_FIELD_VELUE_REPLACEMENTS = "PAR_FIELD_VELUE_REPLACEMENTS";
	public static final String PAR_FIELD_READ_TYPE = "PAR_FIELD_READ_TYPE";
	
	// CSV
	public static final String PAR_CSV_FILENAME = "PAR_CSV_FILENAME";
	public static final String PAR_CSV_DELIMITER = "PAR_CSV_DELIMITER";
	public static final String PAR_CSV_STRING_DELIMITER = "PAR_CSV_STRING_DELIMITER";
	public static final String PAR_CSV_NO_HEADER = "PAR_CSV_NO_HEADER";
	// XML
	public static final String PAR_XML_FILENAME = "PAR_XML_FILENAME";
	public static final String PAR_XML_ROOT_NAME = "PAR_XML_ROOT_NAME";
	public static final String PAR_XML_ELEMENT_NAME = "PAR_XML_ELEMENT_NAME";
	public static final String PAR_XML_ENCODING = "PAR_XML_ENCODING";
	// JSON Export
	public static final String PAR_JSON_FILENAME = "PAR_JSON_FILENAME";
	public static final String PAR_JSON_OBJECT_NAME = "PAR_JSON_OBJECT_NAME";
	
	
	
	
	// Report filter and Title
	public static final String DATE_FROM = "DATE_FROM";
	public static final String DATE_TO = "DATE_TO";
	
	public static final String TITLE_DATE_FROM = "TITLE_DATE_FROM";
	public static final String TITLE_DATE_TO = "TITLE_DATE_TO";
	
	public static final String REPORT_TITLE = "REPORT_TITLE";
	
	// FORMAT_ xxx
	public static final String FORMAT_DATE_HH_MM = "FORMAT_DATE_HH_MM"; // dd.MM.yyyy HH:mm
	public static final String FORMAT_DATE = "FORMAT_DATE"; // dd.MM.yyyy
	public static final String FORMAT_TIME = "FORMAT_TIME"; // HH.mm:ss
	public static final String FORMAT_BACKGROUND_COLORED = "FORMAT_BACKGROUND_COLORED";
	public static final String FORMAT_COLORED = "FORMAT_COLORED";
	public static final String DATE_PATTERN = "DATE_PATTERN'";
	public static final String STYLE_DEZIMAL_1 = "STYLE_DEZIMAL_1";
	public static final String STYLE_DEZIMAL_2 = "STYLE_DEZIMAL_2";
	public static final String STYLE_DEZIMAL_3 = "STYLE_DEZIMAL_3";
	public static final String STYLE_DEZIMAL_4 = "STYLE_DEZIMAL_4";
	public static final String STYLE_DEZIMAL_5 = "STYLE_DEZIMAL_5";
	
	// OPERATOR  Total Fields
	public static final String TOTAL_SUM = "SUM";
	public static final String TOTAL_COUNTER = "COUNTER";
	public static final String TOTAL_AVG = "AVG";
	public static final String TOTAL_MAX = "MAX";
	
	
	// groupSheetHideFields
	public static final String GROUP_SHEET_NAME = "GROUP_SHEET_NAME";
	public static final String GROUP_HIDE_FIELDS = "GROUP_HIDE_FIELDS";
	
	// groupTotalSummarySheet GenericExtendedReport.class
	public static final String PAR_NOT_USE_HIDEFIELDS = "PAR_NOT_USE_HIDEFIELDS";
	
	// optional GenericSummaryModel parameter (GSM_....)
	// GSM .. Parameters is build like an array split by [|] like:  {Block1}|{Block2}|{Block3]...
	// {Field}|{Field}...
	public static final String GSM_GROUP_FIELD = "GSM_GROUP_FIELD";
	//{Field}:[SUM],[COUNTER];{Field}:[SUM]...|{Field}:[SUM,COUNTER];{Field}:[SUM,COUNTER];...
	public static final String GSM_TOTAL_FIELDS = "GSM_TOTAL_FIELDS";
	//{Title}|{Title}...
	public static final String GSM_TITLE = "GSM_TITLE";
	//{Sheetname}|{Sheetname} ...
	public static final String GSM_SHEET_NAME = "GSM_SHEET_NAME";
	// The output Itemrecord result storage   [ItemRecord]:  
	//  _columns = ItemRecord.getMapKeyItems().get("_columns")
	//      data = ItemRecourd.getItems()
	public static final String GSM_RESULT_OBJECT = "GSM_RESULT_OBJECT";
	// {FieldName};{FieldName}|{FieldName} ..
	public static final String GSM_FIELD_ORDER = "GSM_FIELD_ORDER";
	public static final String GSM_TOTAL_TEXT = "GSM_TOTAL_TEXT";
	//{Field}:{Translate};{Field}:{Translate}|...
	public static final String GSM_FIELD_TRANSLATE = "GSM_FIELD_TRANSLATE";
	
	public static final String COL_NAME = "NAME";
	public static final String COL_TYPE = "TYPE";
	
	public static final String REPORT_ID = "REPORT_ID";
	public static final String PAR_TIMEZONE = "PAR_TIMEZONE";
	public static final String REPORT_DATE = "REPORT_DATE";
	
	// connect with mapKeyItem(key) .. as parameter for update, insert, select, delete .. search..
	public static final String PAR_STMT_PLACEHOLDER = "PAR_STMT_PLACEHOLDER";
	public static final String PAR_DATA_RECORD = "PAR_DATA_RECORD";	

}
