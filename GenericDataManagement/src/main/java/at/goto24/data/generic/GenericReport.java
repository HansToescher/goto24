package at.goto24.data.generic;

import java.awt.Color;
import java.io.*;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.TimeZone;

//import at.goto24.report.configuration.xmlbeans.Report;
import at.goto24.data.ComparItems;
import at.goto24.data.ItemRecord;
import at.goto24.data.generic.model.GenericSummaryModel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import org.apache.poi.ss.util.CellRangeAddress;

//import at.goto24.data.util.HSSFUtil;
//import at.goto24.data.util.HSSFUtil.StyleType;
import at.goto24.data.util.FilenameConverter;
import at.goto24.data.util.HSSFUtil;
import at.goto24.data.util.HSSFUtil.StyleType;
import at.goto24.data.util.SqlConstants;
import at.goto24.data.util.Tools;
import at.goto24.data.util.Util;

import org.apache.log4j.Logger;

/**
 * RB-5874 Generic Report
 * Extended by public generic features to use generic rendering methods.
 * Query based Report Generator
 * @author j.toescher
 * @version 4.1.26
 * 
 * It supports Report_TimeZone!
 */
public class GenericReport {
	private static final Logger LOG = Logger.getLogger(GenericReport.class);
	
	private File baseDir = new File(System.getProperty("user.dir") + File.separator + "output" + File.separator + "GenericReport");

	private static final String REPORT_ID = "GenericReport";
	private Connection _con = null;
	private ItemRecord _params = new ItemRecord();
	private Boolean _noDataFound = false;
	
	private ItemRecord _columns = new ItemRecord();
	private ItemRecord _total = new ItemRecord();
	private ItemRecord _group = new ItemRecord();
	
	private SqlConstants  sqlConstants = new SqlConstants();
	
	public GenericReport () {
		super();
	}

	public Collection<File> generateReport(Connection dataSource, ItemRecord argParams)
			throws SQLException, ParseException, IOException {
		Collection<File> returnFiles = new ArrayList<>();
		_con = dataSource;
		
		_noDataFound = false;
		
		if (! baseDir.exists()) {
			baseDir.mkdir();
		}
		
		readParameters(argParams);
		
		ItemRecord _data = readData();
		
		if (! "".equals(_params.getString(Cf.GSM_GROUP_FIELD))) {
			if (_data.getItems().size() > 0) {
				GenericSummaryModel gsm = new GenericSummaryModel(dataSource);
				ItemRecord _gsmData = gsm.readGenericSummaries(_params, _data);
				if (_gsmData != null) {
					_params.addItem(Cf.GSM_RESULT_OBJECT, _gsmData, "ItemRecord");
				}
			}
		}
		
		if (! _noDataFound) {
		
		if ("".equals(_params.getString(Cf.PAR_GROUP_FIELD))) { // Flat Result
		
			if (! "true".equals(_params.getString(Cf.PAR_NO_EXCEL_FILE))) {
				File file = generateExcelSheet(_data);
				if (file != null) {
				returnFiles.add(file);
				}
				
				
			}
			
			if (! "".equals(_params.getString(Cf.PAR_CSV_FILENAME))) {
				File csvfile = generateCSVFile(_data);
				returnFiles.add(csvfile);
			}
			
			if (! "".equals(_params.getString(Cf.PAR_XML_FILENAME))) {
				File xmlfile = generateXMLFile(_data);
				returnFiles.add(xmlfile);
			}
			
			if (! "".equals(_params.getString(Cf.PAR_JSON_FILENAME))) {
				File jsonfile = generateJSONFile(_data);
				returnFiles.add(jsonfile);
			}
		} else { // Group By Result
			prepareGroupData(_data);
			File file =generateGroupExcelSheet(_group, _data);
			returnFiles.add(file);
		}
		}
		
		if (_noDataFound) {
			returnFiles.clear();
		}

		_data = null;
		_total = null;
		_group = null;
		_columns = null;
		
		return returnFiles;
	}
	
	/**
	 * It makes it possible to use the generic report inside of other complex report systems.
	 * @param dataSource as open database connection 
	 * @param argParams ItemRecord as Parameter Container, please use the defined parameter names.
	 * @param resultData ItemRecord is used as data consumer and as return value without printed files!
	 * @return Collection of Files 
	 * @throws SQLException
	 * @throws ParseException
	 * @throws IOException
	 * @throws JRException
	 */
	public Collection<File> generateReport(Connection dataSource, ItemRecord argParams, ItemRecord resultData, boolean readData)
			throws SQLException, ParseException, IOException {
		Collection<File> returnFiles = new ArrayList<File>();
		_con = dataSource;
		_noDataFound = false;
		
		ItemRecord _data = new ItemRecord();
		_params = argParams;
		
		if (_params.getObject(Cf.BASE_DIR) != null) {
			baseDir = (File) _params.getObject(Cf.BASE_DIR);
		}
		
		if (! baseDir.exists()) {
			baseDir.mkdir();
		}
		
		
		if (readData) {
			prepareParsedParameter();
			_data = readData();
		}
		
		if (resultData == null) { // we need only the base result of passed query without printed files!
			if (_data.getItems().size() > 0 ) {
				resultData = _data;
			}
			if (_total.getMapItems().size() > 0 ) {
				resultData.getMapKeyItems().put(Cf.DATA_TOTAL, _total);
			}
			if (_columns.getItems().size() > 0) {
				resultData.getMapKeyItems().put(Cf.DATA_COLUMS, _columns);
			}
		} else {
			_data = resultData;
			_total = resultData.getMapKeyItems().get("_total");
			_columns = resultData.getMapKeyItems().get("_columns");
		}
		
		if (! "".equals(_params.getString(Cf.GSM_GROUP_FIELD))) {
			if (_data.getItems().size() > 0) {
				GenericSummaryModel gsm = new GenericSummaryModel(dataSource);
				ItemRecord _gsmData = gsm.readGenericSummaries(_params, _data);
				if (_gsmData != null) {
					_params.addItem(Cf.GSM_RESULT_OBJECT, _gsmData, "ItemRecord");
				}
			}
		}
		
		if ("".equals(_params.getString(Cf.PAR_GROUP_FIELD))) { // Flat Result
		
			if (! "true".equals(_params.getString(Cf.PAR_NO_EXCEL_FILE))) {
				File file =generateExcelSheet(_data);
				returnFiles.add(file);
			}
			
			if (! "".equals(_params.getString(Cf.PAR_CSV_FILENAME))) {
				File csvfile = generateCSVFile(_data);
				returnFiles.add(csvfile);
			}
			
			if (! "".equals(_params.getString(Cf.PAR_XML_FILENAME))) {
				File xmlfile = generateXMLFile(_data);
				returnFiles.add(xmlfile);
			}
			
			if (! "".equals(_params.getString(Cf.PAR_JSON_FILENAME))) {
				File jsonfile = generateJSONFile(_data);
				returnFiles.add(jsonfile);
			}
			
		} else { // Group By Result
			if ("EXTERNAL_PREPARED".equals(_params.getString(Cf.PAR_GROUP_FIELD))) {
				File file = generateGroupExcelSheetExternalPrepared(_group, _data);
				returnFiles.add(file);
				
			} else {
				prepareGroupData(_data);
				File file =generateGroupExcelSheet(_group, _data);
				returnFiles.add(file);

			}
			
			resultData.getMapKeyItems().put(Cf.DATA_GROUP, _group);
			
		}
		
		if (resultData == null) {
			_data = null;
			_total = null;
			_columns = null;
			_group = null;
		}
		
		return returnFiles;
	}
	
	
	private void readParameters(ItemRecord argParams) throws ParseException, SQLException {
		// dateFrom, dateTo - ' between ? and ? '
		_params = argParams;
		
		String _title = _params.getString(Cf.REPORT_TITLE);
		_title = Tools.getTitle(_title, _params);
		_params.addItem(Cf.REPORT_TITLE, _title, ItemRecord.STRING);
		
		initSql2MySql();
	}	
	

	private void initSql2MySql() {
		if ("".equals(_params.getString(Cf.PAR_SQL_TO_MYSQL))) {
			return;
		}

		Collection<Map<String, String>> _mapList = new ArrayList<>();

		String[] _elements = _params.getString(Cf.PAR_SQL_TO_MYSQL).split("[;]");
		if (_elements.length > 0) {
			for (String _element : _elements) {
				String[] _elems = _element.split("[:]");
				if (_elems.length == 2) {
					String _sqlSeek = _elems[0];
					String _oracleReplace = _elems[1];

					_sqlSeek = regExpressionEscape(_sqlSeek);

					_mapList = sqlConstants.addMySqlReplacements("", _sqlSeek, _oracleReplace, _mapList);

				} else if (_elems.length > 2) {
					String _sqlSeek = _elems[0];

					int x = 0;
					StringBuffer sb = new StringBuffer();
					for (String elem : _elems) {
						if (x > 0) {
							if (x > 1) sb.append(":");
							sb.append(elem);
						}
						x++;
					}
					String _oracleReplace = sb.toString();

					_sqlSeek = regExpressionEscape(_sqlSeek);

					_mapList = sqlConstants.addMySqlReplacements(REPORT_ID, _sqlSeek, _oracleReplace, _mapList);
				}
			}
		}

		_params.addItem(Cf.PAR_SQL_TO_MYSQL, "", ItemRecord.STRING);
	}
	
	private String regExpressionEscape(String arg) {
		String result = "";
		
		StringBuffer sb = new StringBuffer();
		
		for (char _byte : arg.toCharArray()) {
			
			if (_byte == '(') {
				sb.append("\\(");
			} else if (_byte == ')') {
				sb.append("\\)");
			} else if (_byte == '\'') {
				sb.append("\\'");
			} else if (_byte == '+') {
				sb.append("\\+");
			} else {
				sb.append(_byte);
			}
		}
		result = sb.toString();
		
		return result;
	}
	
	
	/**
	 * Prepare the parameter to make it possible for report title and other parser mechanics 
	 * @throws ParseException
	 * @throws SQLException
	 */
	private void prepareParsedParameter() throws ParseException, SQLException {
		
		initSql2MySql();
	}
	

	/**
	 * Consumer interface to get data using query with parameter options
	 * @param dataSource
	 * @param params
	 * @return ItemRecord
	 * @throws SQLException
	 * @throws ParseException
	 */
	public ItemRecord readData(Connection dataSource, ItemRecord params) throws SQLException, ParseException {
		_con = dataSource;
		_params = params;
		prepareParsedParameter();
		return readData();
	}
	
	
	private ItemRecord readData() throws SQLException, ParseException {
		Timestamp tsDateFrom = null;
		Timestamp tsDateTo = null;
		
		if (! "".equals(_params.getString(Cf.PAR_DATE_FROM))) {
			tsDateFrom = new Timestamp(_params.getDate(Cf.DATE_FROM).getTime());
		}
		
		if (! "".equals( _params.getString(Cf.PAR_DATE_TO))) {
			tsDateTo = new Timestamp(_params.getDate(Cf.DATE_TO).getTime());
		}
		
		ItemRecord data = new ItemRecord();
		
		_total = new ItemRecord();
		
		String sQuery = _params.getString(Cf.PAR_QUERY);
		
		sQuery = replaceParams(sQuery);
		
		sQuery = sqlConstants.prepareQuerySyntax(_con, _params.getString(Cf.REPORT_ID), sQuery);
		
		int sx = 1;
		
		try (PreparedStatement stmt = _con.prepareStatement(sQuery)) {
		
		if (tsDateFrom != null && tsDateTo != null ) {
			if (sQuery.toLowerCase().indexOf(" between ? and ? ") > -1) {
				stmt.setTimestamp(sx++, tsDateFrom);
				stmt.setTimestamp(sx++, tsDateTo);
			}
		}
		
		int rx = 1;
		int record = 0;
		
			rx = 1;
			
			try (ResultSet rs =  stmt.executeQuery()) {
			readMetaData(rs);
			
			if (rs == null) {
				_noDataFound = true;
			}
			
			if (rs != null) {
				while (rs.next()) {
					ItemRecord r = new ItemRecord();
					record++;
					for (ItemRecord _col : _columns.getItems()) {
						rx = _col.getInteger("INDEX").intValue();
						String colName = _col.getString("NAME");
						String colType = specialFieldType(colName, _col.getString("ITEMTYPE"));
						
						if (colType.equals(ItemRecord.INTEGER)) {
							r.addItem(colName, rs.getInt(rx), ItemRecord.INTEGER);
						} else if (colType.equals(ItemRecord.DATE2)) {
							if (rs.getTimestamp(rx) != null) {
								Timestamp ts = rs.getTimestamp(rx);
								if (ts != null) {
									r.addItem(colName, new Date(ts.getTime()), ItemRecord.DATE2);
								}
								useTimeZone(r, colName);
							}
						} else if (colType.equals(ItemRecord.STRING)) {
							r.addItem(colName, rs.getString(rx), ItemRecord.STRING);
							
						} else if (colType.equals(ItemRecord.DOUBLE)) {
							r.addItem(colName, rs.getDouble(rx), ItemRecord.DOUBLE);
							
						} else if (colType.equals(ItemRecord.BOOLEAN)) {
							Boolean isWahr =  rs.getBoolean(rx);
							r.addItem(colName, (isWahr ? "Yes" : "No"), ItemRecord.STRING);
						}
						
						computeTotal(r, _col, record, _total);
					}

					data.getItems().add(r);
				}
			}
			
			}
			_noDataFound = data.getItems().isEmpty();
			
			orderBy(data);
		}

		return data;
	}
	
	private void computeTotal(ItemRecord r, ItemRecord _col, int record, ItemRecord _total) {
		if (! "".equals(_params.getString(Cf.PAR_TOTAL_FIELDS))) {
			
			for (String _values : _params.getString(Cf.PAR_TOTAL_FIELDS).split("[;]")) {
				String[] _elems = _values.split("[:]");
				String _fieldName = _elems[0];
				String _operator = _elems[1];
				String colName = _col.getString("NAME");
				String colType = specialFieldType(colName, _col.getString("ITEMTYPE"));
				
				
				if (colName.equals(_fieldName)) {
					if (_operator.equals(Cf.TOTAL_SUM)) {
						ItemRecord.dynamicSummaryOf(_total, r, colName, colType);
					} else if (_operator.equals(Cf.TOTAL_COUNTER)) {
						colType = ItemRecord.INTEGER;
						ItemRecord.dynamicSummaryOf(_total, 1, colName, colType);
					}  else if (_operator.equals(Cf.TOTAL_AVG)) {
						
						if (record == 1) {
							if (colType.equals(ItemRecord.DOUBLE)) {
								Double sum = r.getDouble(_fieldName).doubleValue();
								_total.addItem(_fieldName, sum.doubleValue(), ItemRecord.DOUBLE);
							}
							
							if (colType.equals(ItemRecord.INTEGER)) {
								Integer sum = r.getInteger(_fieldName).intValue();
								_total.addItem(_fieldName, sum.doubleValue(), ItemRecord.DOUBLE);
							}
							
							if (colType.equals(ItemRecord.LONG)) {
								Long sum = r.getLong(_fieldName).longValue();
								_total.addItem(_fieldName, sum.doubleValue(), ItemRecord.DOUBLE);
							}
						} 
						
						if (colType.equals(ItemRecord.DOUBLE)) {
							Double sum = r.getDouble(_fieldName).doubleValue();
							ItemRecord.averageOf(_total, sum.doubleValue(), _fieldName, ItemRecord.DOUBLE);
						}
						
						if (colType.equals(ItemRecord.INTEGER)) {
							Integer sum = r.getInteger(_fieldName).intValue();
							ItemRecord.averageOf(_total, sum.doubleValue(), _fieldName, ItemRecord.DOUBLE);
						}
						
						if (colType.equals(ItemRecord.LONG)) {
							Long sum = r.getLong(_fieldName).longValue();
							ItemRecord.averageOf(_total, sum.doubleValue(), _fieldName, ItemRecord.DOUBLE);
						}
					}
				}
				
			}
		}
	}
	
	private void orderBy(ItemRecord argData) {
		if (argData.getItems().size() > 1) {
			if (! "".equals(_params.getString(Cf.PAR_ORDER_BY))) {
				Collections.sort(argData.getItems(), new ComparItems(_params.getString(Cf.PAR_ORDER_BY)));
			}
		}
	}
	
	private String replaceParams(String query) throws ParseException {
		String result = query;
		
		String _title = _params.getString(Cf.REPORT_TITLE);
		
		if (! "".equals(_params.getString(Cf.PAR_REPLACE_PARAMS))) {
			String _par = _params.getString(Cf.PAR_REPLACE_PARAMS);
			for (String _values : _par.split("[;]")) {
				String[] _elems = _values.split(":");
				if (_elems.length > 1) {
					String _replace = _elems[0];
					String _value = _elems[1];
					String _valueTitle = _elems[1];
					
					if (_value.indexOf("REPORTDATE_") > -1) {
						Date _date = Util.convertString(_value, new Date(), _params.getString(Cf.PAR_TIMEZONE));
						_value = "{ts '" + Tools.getDateFormated(_date, "yyyy-MM-dd HH:mm:ss.S", null) + "'}";
						_valueTitle = Tools.getDateFormated(_date, "yyyy-MM-dd HH:mm:ss.S", null);
					}
					
					result = result.replaceAll(_replace, _value);
				
					_title = _title.replaceAll(_replace, _valueTitle);
				}
			}
			_params.addItem(Cf.REPORT_TITLE, _title, ItemRecord.STRING);
		}
		
		if (! "".equals(_params.getString(Cf.PAR_GROUP_ID))) {
			result = result.replaceFirst("PAR_GROUP_ID", _params.getString(Cf.PAR_GROUP_ID)); 
		}
		
		// replacement dateFrom, dateTo
		Date _date = _params.getDate(Cf.DATE_FROM);
		if (_date != null) {
			String _value = "{ts '" + Tools.getDateFormated(_date, "yyyy-MM-dd HH:mm:ss.S", null) + "'}";
			result = result.replaceAll("DATE_FROM", _value);
		}
		
		_date = _params.getDate(Cf.DATE_TO);
		if (_date != null) {
			String _value = "{ts '" + Tools.getDateFormated(_date, "yyyy-MM-dd HH:mm:ss.S", null) + "'}";
			result = result.replaceAll("DATE_TO", _value);
		}
		
		return result;
	}

	private void readMetaData(ResultSet rs) throws SQLException {
		  int columns =  rs.getMetaData().getColumnCount();
		  _columns = new ItemRecord();
		  
		  if (rs != null) {
		  
		  for (int x = 1; x <= columns; x++) {
			  String colName =  Util.getColumnName(rs.getMetaData(), x);
			  String typeName = rs.getMetaData().getColumnTypeName(x);
			  Integer typeId = rs.getMetaData().getColumnType(x);
			  String typeClassName = rs.getMetaData().getColumnClassName(x);
			  String catalogName = rs.getMetaData().getCatalogName(x);
			  int precision = rs.getMetaData().getPrecision(x);
			  int scale = rs.getMetaData().getScale(x);
			  
			  LOG.debug(">>> colName: " + colName + ", catalogName: " + catalogName + ", type: " + typeName + " (" + precision + "," + scale + "), typeId: " + typeId + ", className: " + typeClassName);
			  
			  if ("true".equals(_params.getString(Cf.PAR_FIELD_TO_UPPERCASE))) {
				  colName = colName.toUpperCase();
			  }
			  
			  ItemRecord _r = new ItemRecord();
			  _r.addItem("NAME", colName, ItemRecord.STRING);
			  _r.addItem("TYPE", typeName, ItemRecord.STRING);
			  _r.addItem("INDEX", x, ItemRecord.INTEGER);

				Util.addItemType(_r, typeName, precision, scale, _params);
			  
			  _columns.getItems().add(_r);
		  }
		  
		  }
	}
	
	private File generateExcelSheet(ItemRecord argData)
			throws IOException, SQLException, ParseException
	{
		File resultFile = null;
		
		Workbook wb = new SXSSFWorkbook();
		CreationHelper createHelper = wb.getCreationHelper();
		
		generateExcel(wb, createHelper, argData);
		
		ItemRecord _gsmData = (ItemRecord) _params.getObject(Cf.GSM_RESULT_OBJECT);
		if (_gsmData != null) {
			for (ItemRecord _par : _gsmData.getItems()) {
				ItemRecord _detailData = (ItemRecord) _par.getObject(Cf.GSM_RESULT_OBJECT);
				if (_detailData != null) {
					ItemRecord _col = _detailData.getMapKeyItems().get("_columns");
					if (_col == null) {
						_col = ItemRecord.createColumnsItem(_detailData.getItems().get(0));
					}
					
					Sheet sheet = wb.createSheet(_par.getString(Cf.PAR_SHEET_TITLE));
					generateGSMDetails(wb, sheet, createHelper, _detailData, _col, _par );
					HSSFUtil.autosizeColumns(sheet, 0, _col.getItems().size() );
				}
			}
			
		}
		
		if ("".equals(_params.getString(Cf.PAR_EXCEL_FILE_NAME))) {
			resultFile = writeXLSFile(baseDir, wb , _params.getDate(Cf.REPORT_DATE));
		} else {
			String resultFilename = baseDir.getAbsolutePath() + File.separator + _params.getString(Cf.PAR_EXCEL_FILE_NAME);
			try (FileOutputStream fileOut = new FileOutputStream(resultFilename))
			{
		        wb.write(fileOut);
		        HSSFUtil.clearWorkbookData(wb);
			}
	        resultFile = new File(resultFilename);
		}
		
		return resultFile;
	}	
	

	

	
	private File generateGroupExcelSheet(ItemRecord argData, ItemRecord detailData)
			throws IOException, SQLException, ParseException
	{
		File resultFile = null;
		
		Workbook wb = new SXSSFWorkbook();
		CreationHelper createHelper = wb.getCreationHelper();
		
		generateGroupExcel(wb, createHelper, argData, detailData);
		
		// GSM GenericSummaryModul print out the results
		ItemRecord _gsmData = (ItemRecord) _params.getObject(Cf.GSM_RESULT_OBJECT);
		if (_gsmData != null) {
			for (ItemRecord _par : _gsmData.getItems()) {
				ItemRecord _detailData = (ItemRecord) _par.getObject(Cf.GSM_RESULT_OBJECT);
				if (_detailData != null) {
					ItemRecord _col = _detailData.getMapKeyItems().get("_columns");
					if (_col == null) {
						_col = ItemRecord.createColumnsItem(_detailData.getItems().get(0));
					}
					
					Sheet sheet = wb.createSheet(_par.getString(Cf.PAR_SHEET_TITLE));
					generateGSMDetails(wb, sheet, createHelper, _detailData, _col, _par );
					HSSFUtil.autosizeColumns(sheet, 0, _col.getItems().size() );
				}
			}
		}
		
		
		if ("".equals(_params.getString(Cf.PAR_EXCEL_FILE_NAME))) {
			resultFile = writeXLSFile(baseDir, wb , _params.getDate(Cf.REPORT_DATE));
		} else {
			String resultFilename = baseDir.getAbsolutePath() + File.separator + _params.getString(Cf.PAR_EXCEL_FILE_NAME);
			try (FileOutputStream fileOut = new FileOutputStream(resultFilename))
			{
		        wb.write(fileOut);
		        HSSFUtil.clearWorkbookData(wb);
			}
	        resultFile = new File(resultFilename);
		}
	
		return resultFile;
	}

	private File generateGroupExcelSheetExternalPrepared(ItemRecord argData, ItemRecord detailData)
			throws IOException, SQLException, ParseException
			{
			File resultFile = null;
			
			Workbook wb = new SXSSFWorkbook();
			CreationHelper createHelper = wb.getCreationHelper();
			
			generateGroupExcelExternalPrepared(wb, createHelper, argData, detailData);
			
			// GSM GenericSummaryModul print out the results
			ItemRecord _gsmData = (ItemRecord) _params.getObject(Cf.GSM_RESULT_OBJECT);
			if (_gsmData != null) {
				for (ItemRecord _par : _gsmData.getItems()) {
					ItemRecord _detailData = (ItemRecord) _par.getObject(Cf.GSM_RESULT_OBJECT);
					if (_detailData != null) {
						ItemRecord _col = _detailData.getMapKeyItems().get("_columns");
						if (_col == null) {
							_col = ItemRecord.createColumnsItem(_detailData.getItems().get(0));
						}
						
						Sheet sheet = wb.createSheet(_par.getString(Cf.PAR_SHEET_TITLE));
						generateGSMDetails(wb, sheet, createHelper, _detailData, _col, _par );
						HSSFUtil.autosizeColumns(sheet, 0, _col.getItems().size() );
					}
				}
			}
			
			
			resultFile = writeXLSFile(baseDir, wb , _params.getDate(Cf.REPORT_DATE));
		
			return resultFile;
	}
	
	private void generateGroupExcel(Workbook wb, CreationHelper createHelper, ItemRecord argData, ItemRecord detailData ) throws SQLException, ParseException {
		int sheetCounter = 0;
		Map<String,String> mapSheets = new HashMap<String, String>();
		String keySheetName = "";
		Sheet sheet;
		
		if ("true".equals(_params.getString(Cf.PAR_GROUP_SHOW_DETAIL_SHEET))) {
			sheet = wb.createSheet(_params.getString(Cf.PAR_SHEET_TITLE));
			generateDetails(wb, sheet, createHelper, detailData);
			HSSFUtil.autosizeColumns(sheet, 0, _columns.getItems().size() );
		}
		
		if ("".equals(_params.getString(Cf.PAR_GROUP_SHEET_ORDER_BY))) {
			for (Entry<String, ItemRecord> _entry : argData.getMapKeyItems().entrySet()) {
				ItemRecord data = _entry.getValue();
				
				if (! "".equals(_params.getString(Cf.PAR_GROUP_RESULT_ORDER_BY))) {
					if (data.getItems().size() > 1) {
						Collections.sort(data.getItems(), new ComparItems(_params.getString(Cf.PAR_GROUP_RESULT_ORDER_BY)));
					}
				}
				
				ItemRecord grTotal = data.getMapKeyItems().get(_params.getString(Cf.PAR_GROUP_FIELD));
				_total = grTotal;
				
				String sheetName = data.getString(_params.getString(Cf.PAR_GROUP_FIELD));
				sheetName = checkedSheetName(sheetName);
				
				sheetCounter++;
				if (sheetName.length() > 25) {
					keySheetName = sheetName.substring(0,25);
				} else {
					keySheetName = sheetName;
				}
				if (mapSheets.get(keySheetName) != null) {
					sheetName = sheetCounter + "_" + sheetName;
				}
				
				_params.addItem(Cf.GROUP_SHEET_NAME, sheetName, ItemRecord.STRING);
				
				sheet = wb.createSheet(sheetName);
				generateDetails(wb, sheet, createHelper, data);
				HSSFUtil.autosizeColumns(sheet, 0, _columns.getItems().size() );
			}
		} else {
			for (Entry<String, ItemRecord> _entry : argData.getMapKeyItems().entrySet()) {
				ItemRecord data = _entry.getValue();
				argData.getItems().add(data);
			}
			
			Collections.sort(argData.getItems(), new ComparItems(_params.getString(Cf.PAR_GROUP_SHEET_ORDER_BY)));
			
			for(ItemRecord data: argData.getItems()) {
				
				if (! "".equals(_params.getString(Cf.PAR_GROUP_RESULT_ORDER_BY))) {
					if (data.getItems().size() > 1) {
						Collections.sort(data.getItems(), new ComparItems(_params.getString(Cf.PAR_GROUP_RESULT_ORDER_BY)));
					}
				}
				
				ItemRecord grTotal = data.getMapKeyItems().get(_params.getString(Cf.PAR_GROUP_FIELD));
				_total = grTotal;
			
				String sheetName = data.getString(_params.getString(Cf.PAR_GROUP_FIELD));
				sheetName = checkedSheetName(sheetName);
				
				sheetCounter++;
				if (sheetName.length() > 25) {
					keySheetName = sheetName.substring(0,25);
				} else {
					keySheetName = sheetName;
				}
				if (mapSheets.get(keySheetName) != null) {
					sheetName = sheetCounter + "_" + sheetName;
				}
				mapSheets.put(keySheetName, sheetName);
				
				_params.addItem(Cf.GROUP_SHEET_NAME, sheetName, ItemRecord.STRING);
				
				sheet = wb.createSheet(sheetName);
				generateDetails(wb, sheet, createHelper, data);
				HSSFUtil.autosizeColumns(sheet, 0, _columns.getItems().size()  - hideFields());
			}
		}
	}	
	
	private void generateGroupExcelExternalPrepared(Workbook wb, CreationHelper createHelper, ItemRecord argData, ItemRecord detailData ) throws SQLException, ParseException {
		
		Sheet sheet;
		
		if ("true".equals(_params.getString(Cf.PAR_GROUP_SHOW_DETAIL_SHEET))) {
			sheet = wb.createSheet(_params.getString(Cf.PAR_SHEET_TITLE));
			_columns =  detailData.getMapKeyItems().get("_columns");
			if (_columns == null) {
				_columns =  ItemRecord.createColumnsItem(detailData.getItems().get(0));
			}
			generateDetails(wb, sheet, createHelper, detailData);
			HSSFUtil.autosizeColumns(sheet, 0, _columns.getItems().size() );
		}
		
		if ("".equals(_params.getString(Cf.PAR_GROUP_SHEET_ORDER_BY))) {
			for (Entry<String, ItemRecord> _entry : argData.getMapKeyItems().entrySet()) {
				ItemRecord data = _entry.getValue();
				
				_columns = data.getMapKeyItems().get("_columns");
				if (_columns == null) {
					_columns = ItemRecord.createColumnsItem(data.getItems().get(0));
				}
				
				if (! "".equals(_params.getString(Cf.PAR_GROUP_RESULT_ORDER_BY))) {
					if (data.getItems().size() > 1) {
						Collections.sort(data.getItems(), new ComparItems(_params.getString(Cf.PAR_GROUP_RESULT_ORDER_BY)));
					}
				}
				
				
				ItemRecord grTotal = data.getMapKeyItems().get("_total");
				_total = grTotal;
				
				
				String sheetName = data.getString("_sheetName");
				
				sheet = wb.createSheet(sheetName);
				generateDetails(wb, sheet, createHelper, data);
				HSSFUtil.autosizeColumns(sheet, 0, _columns.getItems().size() );
			}
		} else {
			
			for (Entry<String, ItemRecord> _entry : argData.getMapKeyItems().entrySet()) {
				ItemRecord data = _entry.getValue();
				argData.getItems().add(data);
			}
			
			Collections.sort(argData.getItems(), new ComparItems(_params.getString(Cf.PAR_GROUP_SHEET_ORDER_BY)));
			
			for(ItemRecord data: argData.getItems()) {
				
				if (! "".equals(_params.getString(Cf.PAR_GROUP_RESULT_ORDER_BY))) {
					if (data.getItems().size() > 1) {
						Collections.sort(data.getItems(), new ComparItems(_params.getString(Cf.PAR_GROUP_RESULT_ORDER_BY)));
					}
				}
				
				_columns = data.getMapKeyItems().get("_columns");
				if (_columns == null) {
					_columns = ItemRecord.createColumnsItem(data.getItems().get(0));
				}
				
				ItemRecord grTotal = data.getMapKeyItems().get("_total");
				_total = grTotal;
			
				String sheetName = data.getString("_sheetName");
				
				sheet = wb.createSheet(sheetName);
				generateDetails(wb, sheet, createHelper, data);
				HSSFUtil.autosizeColumns(sheet, 0, _columns.getItems().size()  - hideFields());
			}
			
		}
		
	}	
	
	private void generateExcel(Workbook wb, CreationHelper createHelper, ItemRecord argData ) throws SQLException, ParseException {
		Sheet sheet = wb.createSheet(_params.getString(Cf.PAR_SHEET_TITLE));
		generateDetails(wb, sheet, createHelper, argData);
		HSSFUtil.autosizeColumns(sheet, 0, _columns.getItems().size() - hideFields() + 1 );
	}
	
	private void generateDetails(Workbook wb, Sheet sheet,
			CreationHelper createHelper, ItemRecord argData) throws SQLException {
		CellStyle defaultStyle = HSSFUtil.getStyle(wb, StyleType.styleLeftBorder);
		CellStyle dateTimeStyle = HSSFUtil.getStyle(wb, StyleType.styleDateTimeBorder);
		CellStyle cellStyleHeader = HSSFUtil.getStyle(wb, StyleType.styleYellowBlackBoldCenter);
		CellStyle numberStyle = HSSFUtil.getStyle(wb, StyleType.styleNumberBorder);
		CellStyle dateTimeSecondsStyle = HSSFUtil.getStyle(wb, StyleType.styleDateTimeWithSecondsBorder);
		CellStyle dateStyle = HSSFUtil.getStyle(wb, StyleType.styleDateBorder);
		CellStyle timeStyle = HSSFUtil.getStyle(wb, StyleType.styleTime2);
		
		Row row = null;
		Cell cell = null;
		
		int _xcol = 0;
		int _xrow = 0;
		
		// Title
		// Header 1
		if (! "".equals(_params.getString(Cf.REPORT_TITLE))) {
			row = sheet.createRow(_xrow);
			cell = row.createCell(_xcol++);
			setCellValue(cell, createHelper, cellStyleHeader,_params.getString(Cf.REPORT_TITLE));
			if (_columns.getItems().size() > 0) {
				sheet.addMergedRegion(new CellRangeAddress(_xrow, _xrow, 0, _columns.getItems().size() - 1 - hideFields()));
			}
			
			// Header 1
			 _xcol = 0;
			 _xrow = 2;
		}
		
		row = sheet.createRow(_xrow++);
		for (ItemRecord header : _columns.getItems()) {
			
			if (isHideField(header.getString("NAME"))) {
				continue;
			}
			
			cell = row.createCell(_xcol++);
			setCellValue(cell, createHelper, cellStyleHeader, convertHeader(header.getString("NAME")));
		}
		
		int _row = _xrow;
		int _col = 0;
		
		// write data
		for (ItemRecord _r : argData.getItems()) {
			_col = 0;
			row = sheet.createRow(_row++);
			
			for (ItemRecord _c : _columns.getItems()) {
				
				if (isHideField(_c.getString("NAME"))) {
					continue;
				}
				
				cell = row.createCell(_col++);
				String itemType = specialFieldType(_c.getString("NAME"), _c.getString("ITEMTYPE"));
				if (itemType.equals(ItemRecord.DATE2)) {
					CellStyle _style = dateTimeSecondsStyle;
					if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE_HH_MM)) {
						_style = dateTimeStyle;
					}
					if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE)) {
						_style = dateStyle;
					}
					if (useFormat(_c.getString("NAME"), Cf.FORMAT_TIME)) {
						_style = timeStyle;
					}
					setCellValue(cell, createHelper, _style, _r.getDate(_c.getString("NAME")));
				} else if (itemType.equals(ItemRecord.INTEGER)) {
					if (useFormat(_c.getString("NAME"), Cf.FORMAT_COLORED)) {
						Color cl = new Color(_r.getInteger(_c.getString("NAME")).intValue(), false);
						String hexColor = String.format("#%02x%02x%02x", cl.getRed(), cl.getGreen(), cl.getBlue());
						setCellValue(cell, createHelper, defaultStyle, hexColor);
					} else {
						if (! "".equals(_params.getString(Cf.PAR_TRANSLATE_FIELD_VALUE))) {
							String _transValue = translateFieldValue(_c.getString("NAME"), ""+_r.getInteger(_c.getString("NAME")));
							if ((""+_r.getInteger("NAME")).equals(_transValue)) {
								setCellValue(cell, createHelper, defaultStyle, _r.getInteger(_c.getString("NAME")));
							} else {
								setCellValue(cell, createHelper, defaultStyle, _transValue);
							}
						} else {
							setCellValue(cell, createHelper, defaultStyle, _r.getInteger(_c.getString("NAME")));
						}
					}
					
				} else if (itemType.equals(ItemRecord.STRING)) {
					setCellValue(cell, createHelper, defaultStyle, translateFieldValue(_c.getString("NAME"), _r.getString(_c.getString("NAME"))));
				}  else if (itemType.equals(ItemRecord.BOOLEAN)) {
					if ("Yes".equals(_r.getString(_c.getString("NAME")))) {
						setCellValue(cell, createHelper, defaultStyle, translateFieldValue(_c.getString("NAME"),"true"));
					} else {
						setCellValue(cell, createHelper, defaultStyle, translateFieldValue(_c.getString("NAME"),"false"));
					}
				}  else if (itemType.equals(ItemRecord.DOUBLE)) {
					CellStyle _style = numberStyle;
					if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_1)) {
						_style = HSSFUtil.getStyle(wb, StyleType.styleNumberDezimal1Border);
					}
					if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_2)) {
						_style = numberStyle;
					}
					if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_3)) {
						_style = HSSFUtil.getStyle(wb, StyleType.styleNumberDezimal3Border);
					}
					if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_4)) {
						_style = HSSFUtil.getStyle(wb, StyleType.styleNumberDezimal4Border);
					}
					
					String _transValue = translateFieldValue(_c.getString("NAME"), ""+_r.getDouble(_c.getString("NAME")));
					if ((""+_r.getDouble(_c.getString("NAME"))).equals(_transValue)) {
						setCellValue(cell, createHelper, _style, _r.getDouble(_c.getString("NAME")));
					} else {
						setCellValue(cell, createHelper, defaultStyle, _transValue);
					}
				}
			}
		}
		
		// Total
		if (! "".equals(_params.getString(Cf.PAR_TOTAL_FIELDS))) {
			ItemRecord _r = _total;
			_col = 0;
			_row++;
			row = sheet.createRow(_row++);
			
			for (ItemRecord _c : _columns.getItems()) {
				if (isHideField(_c.getString("NAME"))) {
					continue;
				}
				
				cell = row.createCell(_col);
				
				if (_col == 0 &&  ! "".equals(_params.getString(Cf.PAR_TOTAL_TEXT)) ) {
					_col++;
					setCellValue(cell, createHelper, defaultStyle, _params.getString(Cf.PAR_TOTAL_TEXT));
				} else {
					_col++;
					String itemType = specialFieldType(_c.getString("NAME"), _c.getString("ITEMTYPE"));
					
					if (isAggregate(_c.getString("NAME"), Cf.TOTAL_AVG)) {
						itemType = ItemRecord.DOUBLE;
					} else if (isAggregate(_c.getString("NAME"), Cf.TOTAL_COUNTER)) {
						itemType = ItemRecord.INTEGER;
					}
					
					if (itemType.equals(ItemRecord.DATE2)) {
						CellStyle _style = dateTimeSecondsStyle;
						if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE_HH_MM)) {
							_style = dateTimeStyle;
						}
						setCellValue(cell, createHelper, _style, _r.getDate(_c.getString("NAME")));
					} else if (itemType.equals(ItemRecord.INTEGER)) {
						setCellValue(cell, createHelper, defaultStyle, _r.getInteger(_c.getString("NAME")));
					} else if (itemType.equals(ItemRecord.STRING)) {
						setCellValue(cell, createHelper, defaultStyle, _r.getString(_c.getString("NAME")));
					}  else if (itemType.equals(ItemRecord.BOOLEAN)) {
						setCellValue(cell, createHelper, defaultStyle, _r.getString(_c.getString("NAME")));
					}  else if (itemType.equals(ItemRecord.DOUBLE)) {
						setCellValue(cell, createHelper, numberStyle, _r.getDouble(_c.getString("NAME")));
					}
				}
			}
		}
	}
	
	private void generateGSMDetails(Workbook wb, Sheet sheet,
			CreationHelper createHelper, ItemRecord argData, ItemRecord _columns, ItemRecord _params) throws SQLException {
		CellStyle defaultStyle = HSSFUtil.getStyle(wb, StyleType.styleLeftBorder);
		CellStyle dateTimeStyle = HSSFUtil.getStyle(wb, StyleType.styleDateTimeBorder);
		CellStyle cellStyleHeader = HSSFUtil.getStyle(wb, StyleType.styleYellowBlackBoldCenter);
		CellStyle numberStyle = HSSFUtil.getStyle(wb, StyleType.styleNumberBorder);
		CellStyle dateTimeSecondsStyle = HSSFUtil.getStyle(wb, StyleType.styleDateTimeWithSecondsBorder);
		CellStyle dateStyle = HSSFUtil.getStyle(wb, StyleType.styleDateBorder);
		CellStyle timeStyle = HSSFUtil.getStyle(wb, StyleType.styleTime2);
		
		Row row = null;
		Cell cell = null;
		
		int _xcol = 0;
		int _xrow = 0;
		
		// Title
		// Header 1
		if (! "".equals(_params.getString(Cf.REPORT_TITLE))) {
			row = sheet.createRow(_xrow);
			cell = row.createCell(_xcol++);
			setCellValue(cell, createHelper, cellStyleHeader,_params.getString(Cf.REPORT_TITLE));
			if (_columns.getItems().size() > 0) {
				sheet.addMergedRegion(new CellRangeAddress(_xrow, _xrow, 0, _columns.getItems().size() - 1 - hideFields()));
			}
			
			// Header 1
			 _xcol = 0;
			 _xrow = 2;
		}
		
		row = sheet.createRow(_xrow++);
		for (ItemRecord header : _columns.getItems()) {
			
			if (isHideField(header.getString("NAME"))) {
				continue;
			}
			
			cell = row.createCell(_xcol++);
			setCellValue(cell, createHelper, cellStyleHeader, convertGSMHeader(header.getString("NAME"), _params));
		}
		
		int _row = _xrow;
		int _col = 0;
		
		// write data
		for (ItemRecord _r : argData.getItems()) {
			_col = 0;
			row = sheet.createRow(_row++);
			
			for (ItemRecord _c : _columns.getItems()) {
				
				if (isHideField(_c.getString("NAME"))) {
					continue;
				}
				
				cell = row.createCell(_col++);
				String itemType = specialFieldType(_c.getString("NAME"), _c.getString("ITEMTYPE"));
				if (itemType.equals(ItemRecord.DATE2)) {
					CellStyle _style = dateTimeSecondsStyle;
					if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE_HH_MM)) {
						_style = dateTimeStyle;
					}
					if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE)) {
						_style = dateStyle;
					}
					if (useFormat(_c.getString("NAME"), Cf.FORMAT_TIME)) {
						_style = timeStyle;
					}
					setCellValue(cell, createHelper, _style, _r.getDate(_c.getString("NAME")));
				} else if (itemType.equals(ItemRecord.INTEGER)) {
					if (useFormat(_c.getString("NAME"), Cf.FORMAT_COLORED)) {
						Color cl = new Color(_r.getInteger(_c.getString("NAME")).intValue(), false);
						String hexColor = String.format("#%02x%02x%02x", cl.getRed(), cl.getGreen(), cl.getBlue());
						setCellValue(cell, createHelper, defaultStyle, hexColor);
					} else {
						if (! "".equals(_params.getString(Cf.PAR_TRANSLATE_FIELD_VALUE))) {
							String _transValue = translateFieldValue(_c.getString("NAME"), ""+_r.getInteger(_c.getString("NAME")));
							if ((""+_r.getInteger("NAME")).equals(_transValue)) {
								setCellValue(cell, createHelper, defaultStyle, _r.getInteger(_c.getString("NAME")));
							} else {
								setCellValue(cell, createHelper, defaultStyle, _transValue);
							}
						} else {
							setCellValue(cell, createHelper, defaultStyle, _r.getInteger(_c.getString("NAME")));
						}
					}
					
				} else if (itemType.equals(ItemRecord.STRING)) {
					setCellValue(cell, createHelper, defaultStyle, translateFieldValue(_c.getString("NAME"), _r.getString(_c.getString("NAME"))));
				}  else if (itemType.equals(ItemRecord.BOOLEAN)) {
					if ("Yes".equals(_r.getString(_c.getString("NAME")))) {
						setCellValue(cell, createHelper, defaultStyle, translateFieldValue(_c.getString("NAME"),"true"));
					} else {
						setCellValue(cell, createHelper, defaultStyle, translateFieldValue(_c.getString("NAME"),"false"));
					}
				}  else if (itemType.equals(ItemRecord.DOUBLE)) {
					CellStyle _style = numberStyle;
					if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_1)) {
						_style = HSSFUtil.getStyle(wb, StyleType.styleNumberDezimal1Border);
					}
					if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_2)) {
						_style = numberStyle;
					}
					if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_3)) {
						_style = HSSFUtil.getStyle(wb, StyleType.styleNumberDezimal3Border);
					}
					if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_4)) {
						_style = HSSFUtil.getStyle(wb, StyleType.styleNumberDezimal4Border);
					}
					
					String _transValue = translateFieldValue(_c.getString("NAME"), ""+_r.getDouble(_c.getString("NAME")));
					if ((""+_r.getDouble(_c.getString("NAME"))).equals(_transValue)) {
						setCellValue(cell, createHelper, _style, _r.getDouble(_c.getString("NAME")));
					} else {
						setCellValue(cell, createHelper, defaultStyle, _transValue);
					}
				}
			}
		}
		
		// Total
		if (! "".equals(_params.getString(Cf.GSM_TOTAL_FIELDS))) {
			ItemRecord _r = _total;
			ItemRecord _xtotal =  argData.getMapKeyItems().get("_total");
			if (_xtotal != null) {
				_r = _xtotal;
			}
			
			_col = 0;
			_row++;
			row = sheet.createRow(_row++);
			
			for (ItemRecord _c : _columns.getItems()) {
				if (isHideField(_c.getString("NAME"))) {
					continue;
				}
				
				cell = row.createCell(_col);
				
				if (_col == 0 &&  ! "".equals(_params.getString(Cf.GSM_TOTAL_TEXT)) ) {
					_col++;
					setCellValue(cell, createHelper, defaultStyle, _params.getString(Cf.GSM_TOTAL_TEXT));
				} else {
					_col++;
					String itemType = specialFieldType(_c.getString("NAME"), _c.getString("ITEMTYPE"));
					
					if (isAggregate(_c.getString("NAME"), Cf.TOTAL_AVG)) {
						itemType = ItemRecord.DOUBLE;
					} else if (isAggregate(_c.getString("NAME"), Cf.TOTAL_COUNTER)) {
						itemType = ItemRecord.INTEGER;
					}
					
					if (itemType.equals(ItemRecord.DATE2)) {
						CellStyle _style = dateTimeSecondsStyle;
						if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE_HH_MM)) {
							_style = dateTimeStyle;
						}
						setCellValue(cell, createHelper, _style, _r.getDate(_c.getString("NAME")));
					} else if (itemType.equals(ItemRecord.INTEGER)) {
						setCellValue(cell, createHelper, defaultStyle, _r.getInteger(_c.getString("NAME")));
					} else if (itemType.equals(ItemRecord.STRING)) {
						setCellValue(cell, createHelper, defaultStyle, _r.getString(_c.getString("NAME")));
					}  else if (itemType.equals(ItemRecord.BOOLEAN)) {
						setCellValue(cell, createHelper, defaultStyle, _r.getString(_c.getString("NAME")));
					}  else if (itemType.equals(ItemRecord.DOUBLE)) {
						setCellValue(cell, createHelper, numberStyle, _r.getDouble(_c.getString("NAME")));
					}
				}
			}
		}
	}	
	
	
	private Boolean isAggregate(String fieldName, String aggregate) {
		Boolean use = false;
		
		if (! "".equals(_params.getString(Cf.PAR_TOTAL_FIELDS))) {
			for (String _values : _params.getString(Cf.PAR_TOTAL_FIELDS).split("[;]")) {
				String[] _elems = _values.split("[:]");
				if (_elems.length > 1) {
					String _fieldName = _elems[0];
					String _aggregate = _elems[1];
					
					if ((fieldName.equalsIgnoreCase(oracleString(_fieldName))) && (aggregate.equals(_aggregate))) {
						use = true;
						break;
					}
				}
			}
		}
		return use;
	}

	private Boolean useFormat(String fieldName, String format) {
		Boolean use = false;
		if (! "".equals(_params.getString(Cf.PAR_FIELD_FORMAT))) {
			for (String _values : _params.getString(Cf.PAR_FIELD_FORMAT).split("[;]")) {
				String[] _elems = _values.split("[:]");
				if (_elems.length > 1) {
					String _fieldName = _elems[0];
					String _format = _elems[1];
					
					if (fieldName.equalsIgnoreCase(oracleString(_fieldName)) && format.equals(_format)) {
						use = true;
						break;
					}
				}
			}
		}
		return use;
	}
	
	private String usePattern(String fieldName, String format) {
		String pattern = "";
		
		if (! "".equals(_params.getString(Cf.PAR_FIELD_FORMAT))) {
			for (String _values : _params.getString(Cf.PAR_FIELD_FORMAT).split("[;]")) {
				String[] _elems = _values.split("[:]");
				if (_elems.length > 1) {
					String _fieldName = _elems[0];
					String _format = _elems[1];
					
					if ((fieldName.equalsIgnoreCase(oracleString(_fieldName))) && (_format.indexOf(_format) > -1)) {
						String[] _xelems = _format.split("'");
						if (_xelems.length > 1) {
							pattern = _xelems[1];
							break;
						}
					}
				}
				
			}
		}
		
		return pattern;
	}

	
	private File generateCSVFile(ItemRecord argData) throws ParseException, FileNotFoundException {
		String csvFilename = baseDir + File.separator + FilenameConverter.convertDateInFileName(_params.getString(Cf.PAR_CSV_FILENAME), new Date());
		
		File file = new File(csvFilename);
		try (PrintWriter writer = new PrintWriter(file)) {
			
			// header
			boolean isFirst = true;
			for (ItemRecord header : _columns.getItems()) {
				
				if (isHideField(header.getString("NAME"))) {
					continue;
				}
				
				if (! isFirst) {
					writer.print(_params.getString(Cf.PAR_CSV_DELIMITER));
				}
				writer.print(translateHeader(header.getString("NAME")));
				
				isFirst = false;
			}
			
			// write data
			for (ItemRecord _r : argData.getItems()) {
				isFirst = true;
				writer.println();
				for (ItemRecord _c : _columns.getItems()) {
					
					if (isHideField(_c.getString("NAME"))) {
						continue;
					}
					
					if (! isFirst) {
						writer.print(_params.getString(Cf.PAR_CSV_DELIMITER));
					}
					
					String itemType = specialFieldType(_c.getString("NAME"), _c.getString("ITEMTYPE"));
					if (itemType.equals(ItemRecord.DATE2)) {
						String _pattern = "yyyy-MM-dd HH:mm:ss";
						
						if (! "".equals(usePattern(_c.getString("NAME"), Cf.DATE_PATTERN))) {
							_pattern = usePattern(_c.getString("NAME"), Cf.DATE_PATTERN); 
						}
						
						if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE_HH_MM)) {
							_pattern = "yyyy-MM-dd HH:mm";
						}
						if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE)) {
							_pattern = "yyyy-MM-dd";
						}
						
						if (_r.getDate(_c.getString("NAME")) != null) {
							writer.print(Tools.getDateFormated(_r.getDate(_c.getString("NAME")), _pattern, null));
						} else {
							writer.print("null");
						}
					} else if (itemType.equals(ItemRecord.INTEGER)) {
						writer.print(_r.getInteger(_c.getString("NAME")));
					} else if (itemType.equals(ItemRecord.STRING)) {
						writer.print(_params.getString(Cf.PAR_CSV_STRING_DELIMITER));
						writer.print(translateFieldValue(_c.getString("NAME"), _r.getString(_c.getString("NAME"))));
						writer.print(_params.getString(Cf.PAR_CSV_STRING_DELIMITER));
					}  else if (itemType.equals(ItemRecord.BOOLEAN)) {
						writer.print(_params.getString(Cf.PAR_CSV_STRING_DELIMITER));
						if ("Yes".equals(_r.getString(_c.getString("NAME")))) {
							writer.print("true");
						} else {
							writer.print("false");
						}
						
						writer.print(_params.getString(Cf.PAR_CSV_STRING_DELIMITER));
					}  else if (itemType.equals(ItemRecord.DOUBLE)) {
						
						String printValue = "" + _r.getDouble(_c.getString("NAME"));  
						
						if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_1)) {
							printValue = String.format("%10.1f", _r.getDouble(_c.getString("NAME"))).trim();
						}
						if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_2)) {
							printValue = String.format("%11.2f", _r.getDouble(_c.getString("NAME"))).trim();
						}
						if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_3)) {
							printValue = String.format("%12.3f", _r.getDouble(_c.getString("NAME"))).trim();
						}
						if (useFormat(_c.getString("NAME"), Cf.STYLE_DEZIMAL_4)) {
							printValue = String.format("%13.4f", _r.getDouble(_c.getString("NAME"))).trim();
						}
						
						writer.print(printValue);
					}
					isFirst = false;
				}
			}
		}
		return file;
	}
		
	private File generateXMLFile(ItemRecord argData) throws ParseException, FileNotFoundException {
		String xmlFilename = baseDir + File.separator + FilenameConverter.convertDateInFileName(_params.getString(Cf.PAR_XML_FILENAME), new Date());
		
		File file = new File(xmlFilename);
		try (PrintWriter writer = new PrintWriter(file)) {
			
			// xml header configuration
			if ("".equals(_params.getString(Cf.PAR_XML_ENCODING))) {
				writer.println("<?xml version=\"1.0\" encoding=\"windows-1250\"?>");
			} else {
				writer.println("<?xml version=\"1.0\" encoding=\"" + _params.getString(Cf.PAR_XML_ENCODING) + "\"?>");
			}
			
			writer.println("<"+_params.getString(Cf.PAR_XML_ROOT_NAME)+">");
			
			// write xml data
			for (ItemRecord _r : argData.getItems()) {
				
				writer.print("   <"+_params.getString(Cf.PAR_XML_ELEMENT_NAME)+" ");
				
				for (ItemRecord _c : _columns.getItems()) {
					
					if (isHideField(_c.getString("NAME"))) {
						continue;
					}
					
					writer.print(translateHeader(_c.getString("NAME"))+"=\"");
					
					String itemType = specialFieldType(_c.getString("NAME"), _c.getString("ITEMTYPE"));
					if (itemType.equals(ItemRecord.DATE2)) {
						String _pattern = "yyyy-MM-dd HH:mm:ss";
						if (! "".equals(usePattern(_c.getString("NAME"), Cf.DATE_PATTERN))) {
							_pattern = usePattern(_c.getString("NAME"), Cf.DATE_PATTERN); 
						}

						if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE_HH_MM)) {
							_pattern = "yyyy-MM-dd HH:mm";
						}
						if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE)) {
							_pattern = "yyyy-MM-dd";
						}
						if (_r.getDate(_c.getString("NAME")) != null) {
							writer.print(Tools.getDateFormated(_r.getDate(_c.getString("NAME")), _pattern, null));
						} else {
							writer.print("null");
						}
					} else if (itemType.equals(ItemRecord.INTEGER)) {
						writer.print(_r.getInteger(_c.getString("NAME")));
					} else if (itemType.equals(ItemRecord.STRING)) {
						writer.print(ToolsEscapeChars.forXML(translateFieldValue(_c.getString("NAME"), _r.getString(_c.getString("NAME")))));
						
					}  else if (itemType.equals(ItemRecord.BOOLEAN)) {
						if ("yes".equals(_r.getString(_c.getString("NAME")))) {
							writer.print("true");
						} else {
							writer.print("false");
						}
					}  else if (itemType.equals(ItemRecord.DOUBLE)) {
						writer.print(_r.getDouble(_c.getString("NAME")));
					}
					writer.print("\" ");
				}
				writer.println("/>");
			}
			writer.print("</"+_params.getString(Cf.PAR_XML_ROOT_NAME)+">");
		}
		return file;
	}	
	
	private File generateJSONFile(ItemRecord argData) throws FileNotFoundException, ParseException {
		String jsonFilename = baseDir + File.separator + FilenameConverter.convertDateInFileName(_params.getString(Cf.PAR_JSON_FILENAME), new Date());
		
		File file = new File(jsonFilename);
		try (PrintWriter writer = new PrintWriter(file)) {
			
			// xml header configuration
			writer.print("{");
			
			// write xml data
			int record = 0;
			int fieldindex = 0;
			for (ItemRecord _r : argData.getItems()) {
				
				if (record > 0) {
					writer.println(",");
				}
				if (! "".equals(_params.getString(Cf.PAR_JSON_OBJECT_NAME))) {
					String _recordId = _params.getString(Cf.PAR_JSON_OBJECT_NAME);
					if (_recordId.indexOf("#") > -1) {
						String _xId = "" + (record + 1);
						_recordId = _recordId.replaceAll("#", _xId);
					}
					writer.print("\""+_recordId+"\":{");
				} 
				
				fieldindex = 0;
				for (ItemRecord _c : _columns.getItems()) {
					if (isHideField(_c.getString("NAME"))) {
						continue;
					}
					
					if (fieldindex > 0) {
						writer.print(",");
					}
					writer.print("\""+translateHeader(_c.getString("NAME"))+"\":");
					
					String itemType = specialFieldType(_c.getString("NAME"), _c.getString("ITEMTYPE"));
					if (itemType.equals(ItemRecord.DATE2)) {
						String _pattern = "yyyy-MM-dd HH:mm:ss";
						if (! "".equals(usePattern(_c.getString("NAME"), Cf.DATE_PATTERN))) {
							_pattern = usePattern(_c.getString("NAME"), Cf.DATE_PATTERN); 
						}

						if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE_HH_MM)) {
							_pattern = "yyyy-MM-dd HH:mm";
						}
						if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE)) {
							_pattern = "yyyy-MM-dd";
						}
						if (_r.getDate(_c.getString("NAME")) != null) {
							writer.print("\"");
							writer.print(Tools.getDateFormated(_r.getDate(_c.getString("NAME")), _pattern, null));
							writer.print("\"");
						} else {
							writer.print("null");
						}
						
					} else if (itemType.equals(ItemRecord.INTEGER)) {
						writer.print(_r.getInteger(_c.getString("NAME")));
					} else if (itemType.equals(ItemRecord.STRING)) {
						writer.print("\"");
						writer.print(ToolsEscapeChars.forJSON(translateFieldValue(_c.getString("NAME"), _r.getString(_c.getString("NAME")))));
						writer.print("\"");
					}  else if (itemType.equals(ItemRecord.BOOLEAN)) {
						if ("Yes".equals(_r.getString(_c.getString("NAME")))) {
							writer.print("true");
						} else {
							writer.print("false");
						}
					}  else if (itemType.equals(ItemRecord.DOUBLE)) {
						writer.print(_r.getDouble(_c.getString("NAME")));
					}
					fieldindex++;
				}
				record++;
				
				if (! "".equals(_params.getString(Cf.PAR_JSON_OBJECT_NAME))) {
					writer.print("}");
				}
			}
			writer.print("}");
		}
		return file;
	}	
	
	private void prepareGroupData(ItemRecord data) throws ParseException {
		_group = new ItemRecord();
		
		
		String _key = "";
		Integer record = 0;
		for (ItemRecord _r : data.getItems()) {
			
			String _value = "";
			String _fieldName = "";
			String _lastColType = "";
			Object _lastValue = null;
			
			for (ItemRecord _c : _columns.getItems()) {
				 _fieldName = _c.getString("NAME");
				 _lastColType = _c.getString("ITEMTYPE");
				String itemType = _c.getString("ITEMTYPE");
				if (itemType.equals(ItemRecord.DATE2)) {
					_lastValue = _r.getDate(_c.getString("NAME"));
					String _pattern = "yyyy-MM-dd HH:mm:ss";
					if (! "".equals(usePattern(_c.getString("NAME"), Cf.DATE_PATTERN))) {
						_pattern = usePattern(_c.getString("NAME"), Cf.DATE_PATTERN); 
					}

					if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE_HH_MM)) {
						_pattern = "yyyy-MM-dd HH:mm";
					}
					if (useFormat(_c.getString("NAME"), Cf.FORMAT_DATE)) {
						_pattern = "yyyy-MM-dd";
					}
					if (_r.getDate(_c.getString("NAME")) != null) {
						_value = Tools.getDateFormated(_r.getDate(_c.getString("NAME")), _pattern, null);
					} else {
						_value = "";
					}
				} else if (itemType.equals(ItemRecord.INTEGER)) {
					_lastValue = _r.getInteger(_c.getString("NAME"));
					_value = ("" + _r.getInteger(_c.getString("NAME")));
				} else if (itemType.equals(ItemRecord.STRING)) {
					_lastValue = _r.getString(_c.getString("NAME"));
					_value = (_r.getString(_c.getString("NAME")));
				}  else if (itemType.equals(ItemRecord.BOOLEAN)) {
					_lastValue = _r.getString(_c.getString("NAME"));
					_value = (_r.getString(_c.getString("NAME")));
				}  else if (itemType.equals(ItemRecord.DOUBLE)) {
					_lastValue = _r.getDouble(_c.getString("NAME"));
					_value = ("" +_r.getDouble(_c.getString("NAME")));
					String _type = specialFieldType(_fieldName, _r.getString("ITEMTYPE"));
					if (ItemRecord.INTEGER.equals(_type)) {
						_value = ("" +_r.getDouble(_c.getString("NAME")).intValue());
					}
				}
	
				if (_fieldName.equalsIgnoreCase(oracleString(_params.getString(Cf.PAR_GROUP_FIELD)))) {
					_key = _value;
					break; // exit
				}
				
			}
			
			ItemRecord _data = _group.getMapKeyItems().get(_key);
			if (_data == null) {
				_data = _r;
				
				String sheetName = _params.getString(Cf.PAR_GROUP_SHEET_PREFIX);
				if (sheetName.indexOf("#") > -1) {
					sheetName = sheetName.replaceAll("#", _value);
				} else {
					sheetName = sheetName + _value;
				}
				
				_data.addItem(_params.getString(Cf.PAR_GROUP_FIELD),
						sheetName, 
						ItemRecord.STRING);
				_data.addItem(_fieldName, _lastValue, _lastColType);
				_group.getMapKeyItems().put(_key, _data);
				record = 1;
			}
			
			ItemRecord _grTotal = _data.getMapKeyItems().get(_params.getString(Cf.PAR_GROUP_FIELD));
			if (_grTotal == null) {
				_grTotal = new ItemRecord();
				_data.getMapKeyItems().put(_params.getString(Cf.PAR_GROUP_FIELD), _grTotal);
			}

			
			for (ItemRecord _col : _columns.getItems()) {
				computeTotal(_r, _col, record, _grTotal);
			}
			
			_data.getItems().add(_r);
			
			record++;
		}
		
		if (! "true".equals(_params.getString(Cf.PAR_GROUP_SHOW_DETAIL_SHEET))) {
			data = new ItemRecord();
		}
		
	}
	
	
	private int hideFields() {
		int result = 0;
		
		if (! "".equals(_params.getString(Cf.PAR_HIDE_FIELDS))) {
			String[] fields = _params.getString(Cf.PAR_HIDE_FIELDS).split("[;]");
			result = fields.length;
		}
		
		return result;
	}
	
	private boolean isHideField(String argFieldName) {
		boolean result = false;
		if (! "".equals(_params.getString(Cf.PAR_HIDE_FIELDS))) {
			String[] fields = _params.getString(Cf.PAR_HIDE_FIELDS).split("[;]");
			for (String _field : fields) {
				if (oracleString(_field).equals(argFieldName)) {
					result = true;
					break;
				}
			}
		}
		
		if (! "".equals(_params.getString(Cf.PAR_GROUP_SHEET_HIDE_FIELDS))) {
			String[] hideElems = _params.getString(Cf.PAR_GROUP_SHEET_HIDE_FIELDS).split("[;]");
			for (String _grpElem : hideElems) {
				String[] a_GrpElems = _grpElem.split("[:]");
				String grpValue = _params.getString(Cf.GROUP_SHEET_NAME);
				if (grpValue.equals(a_GrpElems[0])) {
					for(int x = 1; x < a_GrpElems.length; x++) {
						String _field = a_GrpElems[x];
						if (oracleString(_field).equals(argFieldName)) {
							result = true;
							break;
						}
					}
				}
			}
		}
		
		return result;
	}
	
	
	private String convertGSMHeader(String header, ItemRecord _param) {
		String result = header;
		if (! "".equals(_param.getString(Cf.GSM_FIELD_TRANSLATE).trim())) {
		String[] aElems = _param.getString(Cf.GSM_FIELD_TRANSLATE).split("[;]");
		if (aElems.length > 0) {
			for(String elem : aElems) {
				String[] afield = elem.split("[:]");
				if (afield.length > 0) {
					String field = afield[0];
					String translate = afield[1];
					if (field.equals(header)) {
						result = translate;
						break;
					}
				}
			}
		}
		}
		
		return result;
	}
	
	private String convertHeader(String header) {
		String result = header;
		
		result = translateHeader(header);
		
		result = result.replaceAll("_", " ");
		
		return result;
	}
	
	private String specialFieldType(String fieldName, String orgType) {
		String result = orgType;
		
		String[] fields = _params.getString(Cf.PAR_TYPE_OF_FIELD).split("[;]");
		for (String valuePears: fields) {
			String[] elems = valuePears.split("[:]");
			if (elems.length > 1) {
				if (fieldName.equalsIgnoreCase(oracleString(elems[0]))) {
					result = elems[1];
					break;
				}
			}
		}
		
		
		return result;
	}
	
	
	private String oracleString(String value) {
		String result = value;
		if (Util.isOracleConnection(_con)) {
			result = value.toUpperCase();
		}
		return result;
	}
	
	private String translateHeader(String value) {
		String result = value;
		if (! "".equals(_params.getString(Cf.PAR_TRANSLATE_HEADER))) {
			String[] translates = _params.getString(Cf.PAR_TRANSLATE_HEADER).split("[;]");
			for (String element : translates) {
				String[] elems = element.split("[:]");
				if (elems.length > 1) {
					String _fieldName = oracleString(elems[0]);
					String _translate = elems[1];
					if (value.toUpperCase().equals(_fieldName.toUpperCase())) {
						result = _translate;
						break;
					}
				}
				
			}
		}
		
		
		return result;
	}
	
	
	/**
	 * It is only usable for String values.
	 * e.g. "true" -&gt; 'correct'
	 * @param fieldname
	 * @param value
	 * @return String
	 */
	private String translateFieldValue(String fieldname, String value) {
		String result = value;
		if (! "".equals(_params.getString(Cf.PAR_TRANSLATE_FIELD_VALUE))) {
			String[] elements = _params.getString(Cf.PAR_TRANSLATE_FIELD_VALUE).split("[;]");
			for (String element : elements) {
				String[] _elems = element.split("[:]");
				if (_elems.length > 2) {
					String _field = _elems[0];
					String _value = _elems[1];
					String _newValue = _elems[2];
					
					
					
					if (fieldname.toUpperCase().equals(_field.toUpperCase())) {
						if ("CONVERT_OFFSET".equals(_newValue)) {
							TimeZone tz = TimeZone.getTimeZone(value);
							if (tz != null) {
								int offset = tz.getRawOffset();
								offset = offset / 1000;
								offset = offset / 3600;
								
								String _vz = "";
								if (offset >= 0 ) {
									_vz = "+";
								}
								String _offset = "GMT" + _vz + offset;
								_newValue = _offset;
								result = _newValue;
								break;
							}
						}
						
						if (value.equals(_value)) {
							result = _newValue;
							break;
						}
					}
					
					
				}
			}
		}
		return result;
	}
	
	
	private String checkedSheetName(String name) {
		String result = "";
		StringBuffer sb = new StringBuffer();
		int x = 0;
		for (char _char : name.toCharArray()) {
			if ("/".indexOf(_char) < 0) {	
				sb.append(_char);
			} else if ("/".indexOf(_char) >= 0) {
				sb.append('_');
			}
			
		}
		result = sb.toString();
		if (x > 1) {
			result = sb.toString() + "_" + x;
		}
		
		return result;
	}
	
	
	
	/**
	 * Recalculate the date time value depending on time zone (JodaTime)
	 * @param r ItemRecord (raw data by db result)
	 * @param fieldName String
	 * @return boolean 
	 */
	private Boolean useTimeZone(ItemRecord r, String fieldName) {
		Boolean ok = false;
		
		if (! "".equals(_params.getString(Cf.PAR_FIELD_TIMEZONE))) {
			String[] a_fieldparams = _params.getString(Cf.PAR_FIELD_TIMEZONE).split("[;]");
			if (a_fieldparams.length > 0) {
				for (String _fieldparam : a_fieldparams) {
					String[] a_fp = _fieldparam.split("[:]");
					if (a_fp.length == 3) {
						String _fieldName = a_fp[0];
						String _type = a_fp[1];
						String _timeZone = a_fp[2];
						String _parOffset = a_fp[2];
						
						if (_fieldName.toUpperCase().equals(fieldName.toUpperCase())) {
							if ("TIMEZONE".equals(_type)) {
								TimeZone _tz = TimeZone.getTimeZone(_timeZone);
								Integer _offset =  _tz.getRawOffset();
								Long offsetHour =  _offset.longValue();
								Long _timeOfDate = r.getDate(fieldName).getTime() + offsetHour;
								Date _date = new Date(_timeOfDate.longValue());
								r.addItem(fieldName, _date, ItemRecord.DATE2);
								ok = true;
							} else if ("OFFSET".equals(_type)) {
								int _offset =  Integer.parseInt(_parOffset);
								Long offsetHour = 1000L * 60L * 60L;
								offsetHour = offsetHour * _offset;
								Long _timeOfDate = r.getDate(fieldName).getTime() + offsetHour;
								Date _date = new Date(_timeOfDate.longValue());
								r.addItem(fieldName, _date, ItemRecord.DATE2);
								ok = true;
							} else if ("FIELDNAME".equals(_type)) {
								String r_timeZone = r.getString(_timeZone);
								TimeZone _tz = TimeZone.getTimeZone(r_timeZone);
								Integer _offset =  _tz.getRawOffset();
								Long offsetHour =  _offset.longValue();
								Long _timeOfDate = r.getDate(fieldName).getTime() + offsetHour;
								Date _date = new Date(_timeOfDate.longValue());
								r.addItem(fieldName, _date, ItemRecord.DATE2);
								ok = true;
							}
							
							break;
						}
						
					}
				}
			}
			
		}
		
		return ok;
	}
	
	public File writeXLSFile(File baseDir, Workbook wb, Date date) throws IOException {
        return writeXLSFile(baseDir, wb, date, null);
    }

    public File writeXLSFile(File baseDir, Workbook wb, Date date,
    		Map<String, String> placeholderReplacementMap) throws IOException {
    	return writeXLSFile(baseDir, wb, date, placeholderReplacementMap, null);
    }	
    
    /**
     * RB-845 Report ID is a part of Report Result File Name
     * @param baseDir
     * @param wb
     * @param date
     * @param placeholderReplacementMap
     * @param filename
     * @return
     * @throws IOException
     */
    //writeXLSFile with Filename param
    public File writeXLSFile(File baseDir, Workbook wb, Date date, Map<String, String> placeholderReplacementMap, String filename) throws IOException {
    	
    	String resultFilename = null;
    	if (!baseDir.exists()) {
            baseDir.mkdirs();
        }
    	if (filename != null) {
    		resultFilename = baseDir + File.separator + FilenameConverter.convertDateInFileName(filename,date);
    	} else {
    		resultFilename = baseDir + File.separator + FilenameConverter.convertDateInFileName(_params.getString(Cf.PAR_EXCEL_FILE_NAME), date);
    	}
    	
        resultFilename = FilenameConverter.replacePlaceholdersInFilename(resultFilename, placeholderReplacementMap); //replace all properties used as placeholder in config filename
        FileOutputStream fileOut = new FileOutputStream(resultFilename);
        wb.write(fileOut);
        HSSFUtil.clearWorkbookData(wb);
        fileOut.close();
    	return new File(resultFilename);
    }  
    
    protected void setCellValue (Cell cell, CreationHelper createHelper, String value)
    {
    	setCellValue (cell, createHelper, null, value);
    }
    
    protected void setCellValue (Cell cell, CreationHelper createHelper,
    		CellStyle cellStyle, String value)
    {
    	if (cellStyle!=null)
    	{
    		cell.setCellStyle(cellStyle);
    	}
    	if (value!=null)
    	{
    		cell.setCellValue(createHelper.createRichTextString(value));
    	}
    }
    
    protected void setCellValue (Cell cell, CreationHelper createHelper,
    		Date value)
    {
    	setCellValue (cell, createHelper, null, value);
    }
    
    protected void setCellValue (Cell cell, CreationHelper createHelper,
    		CellStyle cellStyle, Date value)
    {
    	if (cellStyle!=null)
    	{
    		cell.setCellStyle(cellStyle);
    	}
    	if (value!=null)
    	{
    		cell.setCellValue(value);
    	}
    }
    
    /**Method to set a numeric cell value.
     * Note that also non-double numeric values can be set with this method (e.g. int values, ...)
     * 
     * @param cell
     * @param createHelper
     * @param value
     */
    protected void setCellValue (Cell cell, CreationHelper createHelper,
    		Number value)
    {
    	setCellValue (cell, createHelper, null, value);
    }
    
    /**Method to set a numeric cell value.
     * 
     * @param cell
     * @param createHelper
     * @param cellStyle
     * @param value
     */
    protected void setCellValue (Cell cell, CreationHelper createHelper,
    		CellStyle cellStyle, Number value)
    {
    	if (cellStyle!=null)
    	{
    		cell.setCellStyle(cellStyle);
    	}
    	if (value!=null)
    	{
    		cell.setCellValue(value.doubleValue());
    	}
    }    
}
