package at.goto24.data.manager;

import java.beans.PropertyVetoException;
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
import java.util.Map;
import java.util.Map.Entry;

//import at.goto24.report.report.AbstractReport;
import at.goto24.data.util.Util;
import org.apache.log4j.Logger;

import com.mchange.v2.c3p0.ComboPooledDataSource;
import at.goto24.data.ComparItems;
import at.goto24.data.ItemRecord;
import at.goto24.data.generic.Cf;
import at.goto24.data.util.SqlConstants;
import at.goto24.data.util.Tools;

import static at.goto24.data.util.Util.ITEMTYPE;
import static at.goto24.data.util.Util.NAME;


/**
 * DataManager:
 * It opens a connection and you can read data by parameter definition!
 * The result is an ItemRecord type with column description '_columns'  and optional with summaries '_total'
 * @author j.toescher
 *
 */
public class DataManager {
	private static final Logger LOG = Logger.getLogger(DataManager.class);

	public static final String REPORTID = "DataManager";
	private Connection con = null;
	
	// Connection Item Names
	public static final String DriverClass = "DriverClass";
	public static final String JdbcUrl = "JdbcUrl";
	public static final String ConnectionTestPeriod = "ConnectionTestPeriod";
	public static final String PreferredTestQuery = "PreferredTestQuery";
	
	
	/** column description types **/
	private static final String INDEX = "INDEX";
	private static final String TYPE = "TYPE";
	
	
	/**
	 * Create parameter container
	 * @param driverClass
	 * @param jdbcUrl
	 * @return ItemRecord as Connection Parameter
	 */
	public static ItemRecord CreateParam(String driverClass, String jdbcUrl ) {
		ItemRecord param = new ItemRecord();
		param.addItem(DriverClass, driverClass, ItemRecord.STRING);
		param.addItem(JdbcUrl, jdbcUrl, ItemRecord.STRING);
		return param;
	}
	
	public DataManager() {
	}
	
	public DataManager(Connection dataSource) {
		con = dataSource;
	}
	
	public DataManager(ItemRecord argParams) throws SQLException, PropertyVetoException {
		doOpen(argParams);
	}
	
	public void close() throws SQLException {
		if (con != null) {
			con.close();
		}
	}
	
	
	public Connection getConnection() {
		return con;
	}
	
	
	/**
	 * 
	 * @param argParams Parameters to open the database
	 * @return boolean
	 */
	public boolean doOpen(ItemRecord argParams) throws SQLException, PropertyVetoException {
		boolean ok = false;
		ItemRecord params = argParams;
		ComboPooledDataSource cpds = new ComboPooledDataSource();


			cpds.setDriverClass(params.getString(DriverClass)); // loads the jdbc driver
			cpds.setJdbcUrl(params.getString(JdbcUrl));
			Integer testPeriod = 0;
			if (! "".equals(params.getString(ConnectionTestPeriod))) {
				testPeriod = Integer.valueOf(params.getString(ConnectionTestPeriod));
			}
			if ( 0 != testPeriod.intValue()) {
				cpds.setIdleConnectionTestPeriod(testPeriod);
				cpds.setMaxConnectionAge(testPeriod);
				cpds.setTestConnectionOnCheckout(true);
				if (! "".equals(params.getString(PreferredTestQuery))) {
					cpds.setPreferredTestQuery(params.getString(PreferredTestQuery));
					//use own test query to check if connection is alive (otherwise csv sources won't set up)
				}
			}

		  con = cpds.getConnection();

			if (con != null) {
				ok = true;
			}

		return ok;
	}

	/**
	 * Read Data by Query Passing
	 * @param report
	 * @param argQuery String
	 * @return Data placed in ItemRecord.getItems() 
	 * @throws SQLException
	 * @throws ParseException
	 */
	public ItemRecord readData(String argQuery) throws SQLException, ParseException  {
		ItemRecord dbparams = new ItemRecord();
		dbparams.addItem(Cf.PAR_QUERY, argQuery, ItemRecord.STRING);
		return readData(dbparams);
	}
	
	
	/**
	 * Parameters is ItemRecord to define query, order and summaries.
	 * @param report
	 * @param _params 
	 * @return ItemRecord include _columns, _total
	 * @throws SQLException
	 * @throws ParseException
	 */
	public ItemRecord readData(ItemRecord params) throws SQLException, ParseException {
		Timestamp tsDateFrom = null;
		Timestamp tsDateTo = null;
		
		if (! "".equals(params.getString(Cf.PAR_DATE_FROM))) {
			tsDateFrom = new Timestamp(params.getDate(Cf.DATE_FROM).getTime());
		}
		
		if (! "".equals( params.getString(Cf.PAR_DATE_TO))) {
			tsDateTo = new Timestamp(params.getDate(Cf.DATE_TO).getTime());
		}


		SqlConstants sqlConstants = null;
		sqlConstants = new SqlConstants();
		
		
		initSql2Oracle(sqlConstants, params);
		initSql2MySql(sqlConstants, params);
		
		ItemRecord data = new ItemRecord();
		ItemRecord total = new ItemRecord();
		String sQuery = params.getString(Cf.PAR_QUERY);
		sQuery = replaceParams(sQuery, params);

		String reportName = params.getString("PAR_REPORTNAME");
		if (reportName != null && ! "".equals(reportName)) {
			sQuery = sqlConstants.prepareQuerySyntax(con, reportName, sQuery);
		} else {
			sQuery = sqlConstants.prepareQuerySyntax(con, REPORTID, sQuery);
		}
		
		int sx = 1;
		try  (PreparedStatement stmt = con.prepareStatement(sQuery)) {
		
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
		
		if (rs != null) {
			ItemRecord columns = readMetaData(rs, params);
			data.getMapKeyItems().put("_columns", columns);
			
			while (rs.next()) {
				ItemRecord r = new ItemRecord();
				record++;
				for (ItemRecord col : columns.getItems()) {
					rx = col.getInteger(INDEX).intValue();
					String colName = col.getString(NAME);
					String colType = col.getString(ITEMTYPE);
					
					if (colType.equals(ItemRecord.INTEGER)) {
						r.addItem(colName, rs.getInt(rx), ItemRecord.INTEGER);
					} else if (colType.equals(ItemRecord.DATE2)) {
						if (rs.getTimestamp(rx) != null) {
							Timestamp ts = rs.getTimestamp(rx);
							if (ts != null) {
								r.addItem(colName, new Date(ts.getTime()), ItemRecord.DATE2);
							}
						}
					} else if (colType.equals(ItemRecord.STRING)) {
						r.addItem(colName, rs.getString(rx), ItemRecord.STRING);
						
					} else if (colType.equals(ItemRecord.DOUBLE)) {
						r.addItem(colName, rs.getDouble(rx), ItemRecord.DOUBLE);
						
					} else if (colType.equals(ItemRecord.BOOLEAN)) {
						Boolean isWahr =  rs.getBoolean(rx);
						r.addItem(colName, (isWahr ? "Yes" : "No"), ItemRecord.STRING);
					}
					
					computeTotal(r, col, record, total, params);
				}
				
				data.getItems().add(r);
			}
		}
		}
		
		data.getMapKeyItems().put("_total", total);
		
		orderBy(data, params);
		
		}
		return data;
	}
	
	private ItemRecord readMetaData(ResultSet rs, ItemRecord _params) throws SQLException {
		  int columns =  rs.getMetaData().getColumnCount();
		  ItemRecord _columns = new ItemRecord();
		  
		  if (rs != null) {
		  
		  for (int x = 1; x <= columns; x++) {
				String colName =  Util.getColumnName(rs.getMetaData(), x);
			  String typeName = rs.getMetaData().getColumnTypeName(x);
			  Integer typeId = rs.getMetaData().getColumnType(x);
			  String typeClassName = rs.getMetaData().getColumnClassName(x);
			  String catalogName = rs.getMetaData().getCatalogName(x);
			  int precision = rs.getMetaData().getPrecision(x);
				int scale = rs.getMetaData().getScale(x);
			  
			  LOG.debug(">>> colName: " + colName + ", catalogName: " + catalogName + ", type: " + typeName + ", typeId: " + typeId + ", className: " + typeClassName);
			  
			  ItemRecord _r = new ItemRecord();
			  _r.addItem(NAME, colName, ItemRecord.STRING);
			  _r.addItem(TYPE, typeName, ItemRecord.STRING);
			  _r.addItem(INDEX, x, ItemRecord.INTEGER);

				Util.addItemType(_r, typeName, precision, scale, _params);
			  
			  _columns.getItems().add(_r);
		  }
		  
		  }
		  return _columns;
	}

	private void computeTotal(ItemRecord r, ItemRecord _col, int record, ItemRecord _total, ItemRecord _params) {
		if (! "".equals(_params.getString(Cf.PAR_TOTAL_FIELDS))) {
			
			for (String _values : _params.getString(Cf.PAR_TOTAL_FIELDS).split("[;]")) {
				String[] _elems = _values.split("[:]");
				String _fieldName = _elems[0];
				String _operator = _elems[1];
				String colName = _col.getString(NAME);
				String colType = _col.getString(ITEMTYPE);
				
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
	
	private void orderBy(ItemRecord argData, ItemRecord _params) {
		if (argData.getItems().size() > 1) {
			if (! "".equals(_params.getString(Cf.PAR_ORDER_BY))) {
				Collections.sort(argData.getItems(), new ComparItems(_params.getString(Cf.PAR_ORDER_BY)));
			}
		}
	}
	
	private String replaceParams(String query, ItemRecord _params) throws ParseException {
		String result = query;
		
		if (! "".equals(_params.getString(Cf.PAR_REPLACE_PARAMS))) {
			String _par = _params.getString(Cf.PAR_REPLACE_PARAMS);
			for (String _values : _par.split("[;]")) {
				String[] _elems = _values.split(":");
				if (_elems.length > 1) {
					String _replace = _elems[0];
					String _value = _elems[1];
					
					result = result.replaceAll(_replace, _value);
				
				}
			}
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
	
	public boolean isOracle(Connection con) {
    	return (con.toString().indexOf("oracle") > -1 );
    }
	
	
	private void initSql2Oracle(SqlConstants sqlConstants, ItemRecord _params) {
		if ("".equals(_params.getString(Cf.PAR_SQL_TO_ORACLE))) {
			return;
		}
		
		Collection<Map<String, String>> _mapList = new ArrayList<>();
		
		String[] _elements = _params.getString(Cf.PAR_SQL_TO_ORACLE).split("[;]");
		if (_elements.length > 0) {
			for (String _element : _elements) {
				String[] _elems = _element.split("[:]");
				if (_elems.length == 2) {
					String _sqlSeek = _elems[0];
					String _oracleReplace = _elems[1];
					
					_sqlSeek = regExpressionEscape(_sqlSeek);
					
					_mapList = sqlConstants.addOracleReplacements(REPORTID, _sqlSeek, _oracleReplace, _mapList);
					
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
					
					_mapList = sqlConstants.addOracleReplacements(REPORTID, _sqlSeek, _oracleReplace, _mapList);
				}
			}
		}
		_params.addItem(Cf.PAR_SQL_TO_ORACLE, "", ItemRecord.STRING);
	}

	private void initSql2MySql(SqlConstants sqlConstants, ItemRecord _params) {
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

					_mapList = sqlConstants.addMySqlReplacements(REPORTID, _sqlSeek, _oracleReplace, _mapList);

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

					_mapList = sqlConstants.addMySqlReplacements(REPORTID, _sqlSeek, _oracleReplace, _mapList);
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
	 * To execute data update by sql statements
	 * @param params - well formed ItemRecord structure
	 * @return counter of updated records
	 */
	public Integer executeDataUpdate(ItemRecord params) {
		Integer result = -1;
		PreparedStatement stmt = null;
		try {
			String qry = params.getString(Cf.PAR_QUERY);
			
			qry = replaceParams(qry, params);
			
			qry = fillQuery(qry, params);
			
			con.setAutoCommit(false);
			
			stmt = con.prepareStatement(qry);
			
			result = stmt.executeUpdate();
			if (result > 0) {
				con.commit();
			} else {
				con.rollback();
			}
		} catch (Exception ex) {
			try {
				con.rollback();
			} catch (Exception e) {
				LOG.error(ex.getMessage(), ex);
			}
			LOG.error(ex.getMessage(), ex);
		} finally {
			if (stmt != null) {
				try {
					stmt.close();
				} catch (Exception ex ) {
					LOG.error(ex.getMessage(), ex);
				}
			}
		}
		return result;
	}	
	
	/**
	 * Used for update, insert and delete query .. called by execute()
	 * @param qry String
	 * @param param ItemRecord as Parameter with Data ItemRecord
	 * @return
	 */
	private String fillQuery(String qry, ItemRecord param) {
		String result = qry;
		try {
			
			if ("".equals(param.getString(Cf.PAR_STMT_PLACEHOLDER))) {
				return result;
			}
			
						
			ItemRecord rec = param.getMapKeyItems().get(Cf.PAR_DATA_RECORD);
			if (rec == null) {
				return result;
			}
			
			String q = qry;
			String[] placeholder = param.getString(Cf.PAR_STMT_PLACEHOLDER).split("[;]");
			
			for(String elems : placeholder) {
				String[] elem = elems.split("[:]");
				String place = elem[0];
				String field = elem[1];
				String value = "";
				
				for(Entry<String, ItemRecord> entry : rec.getMapItems().entrySet()) {
					ItemRecord f = entry.getValue();
					if (f.getName().toUpperCase().equals(field.toUpperCase())) {
						if (ItemRecord.BOOLEAN.equals(f.getClassName())) {
							value = "" + (f.getCheck()?"1":"0") ;
						} else if (ItemRecord.STRING.equals(f.getClassName())) {
							value = "'" + f.getText() + "'";
						} else if (ItemRecord.INTEGER.equals(f.getClassName())) {
							value = "" + f.getIntValue();
						} else if (ItemRecord.DOUBLE.equals(f.getClassName())) {
							value = "" + f.getDoubleValue();
						} else if (ItemRecord.DATE2.equals(f.getClassName())) {
							value = "{ts '" + Tools.getDateFormated(f.getDate(), "yyyy-MM-dd HH:mm:ss", null) + "'}";
						}
					}
				}
				
				q = q.replaceAll(place, value);
			}
			
			result = q;
			
			
		} catch (Exception ex) {
			LOG.error(ex.getMessage(), ex);
		}
		
		// debug
		System.out.println(">>> execute Query: [" + result + "]");
		
		return result;
	}	
		
}
