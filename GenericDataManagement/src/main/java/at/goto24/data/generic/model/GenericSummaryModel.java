package at.goto24.data.generic.model;

import java.sql.Connection;
import java.util.Map.Entry;

import org.apache.log4j.Logger;

import at.goto24.data.ItemRecord;
import at.goto24.data.generic.Cf;
import at.goto24.data.util.Util;

public class GenericSummaryModel {
	private static final Logger LOG = Logger.getLogger(GenericSummaryModel.class);
	
	private Connection _con;
	
	public GenericSummaryModel(Connection dataSource) {
		super();
		_con = dataSource;
	}
	
	
	public ItemRecord readGenericSummaries(ItemRecord _param, ItemRecord _detailData) {
		ItemRecord result = null;
		try {
			result = prepareParamArray(_param);
			for (ItemRecord _par: result.getItems()) {
				ItemRecord _gsmData = dynamicFieldGroup(_detailData, _par);
				_par.addItem(Cf.GSM_RESULT_OBJECT, _gsmData, "ItemRecord");
				
				// FieldOrder...
				ItemRecord _col = _gsmData.getMapKeyItems().get("_columns");
				ItemRecord _colNew = new ItemRecord();
				String[] aFieldOrder = _par.getString(Cf.GSM_FIELD_ORDER).split("[;]");
				if (aFieldOrder.length > 0) {
					for (String orderField : aFieldOrder) {
						for (ItemRecord _c : _col.getItems()) {
							if (orderField.equals(_c.getString("NAME"))) {
								_colNew.getItems().add(_c);
							}
						}
					}
					if (_colNew.getItems().size() > 0) {
						_gsmData.getMapKeyItems().put("_columns", _colNew);
					}
				}
				
			}
			
		} catch (Exception ex) {
			LOG.error(ex.getMessage(), ex);
		}
		return result;
	}
	
	
	public ItemRecord prepareParamArray(ItemRecord _params) {
		ItemRecord _result = new ItemRecord();
		try {
			if (! "".equals(_params.getString(Cf.GSM_GROUP_FIELD))) {
				String[] aGroupFields = _params.getString(Cf.GSM_GROUP_FIELD).split("[|]");
				String[] aTotalFields = _params.getString(Cf.GSM_TOTAL_FIELDS).split("[|]");
				String[] aSheetNames = _params.getString(Cf.GSM_SHEET_NAME).split("[|]");
				String[] aTitles = _params.getString(Cf.GSM_TITLE).split("[|]");
				String[] aFieldOrder = _params.getString(Cf.GSM_FIELD_ORDER).split("[|]");
				String[] aTotalText = _params.getString(Cf.GSM_TOTAL_TEXT).split("[|]");
				String[] aFieldTranslate = _params.getString(Cf.GSM_FIELD_TRANSLATE).split("[|]");

				if (aGroupFields.length > 0) {
					for (int x = 0; x < aGroupFields.length; x++) {
						ItemRecord _r = new ItemRecord();
						_r.addString(Cf.GSM_GROUP_FIELD, aGroupFields[x]);
						_r.addString(Cf.GSM_TOTAL_FIELDS, aTotalFields[x]);
						_r.addString(Cf.GSM_SHEET_NAME, aSheetNames[x]);
						_r.addString(Cf.GSM_TITLE, aTitles[x]);
						_r.addString(Cf.PAR_SHEET_TITLE, aTitles[x]);
						_r.addString(Cf.REPORT_TITLE, aTitles[x]);
						_r.addString(Cf.GSM_FIELD_ORDER, aFieldOrder[x]);
						_r.addString(Cf.GSM_TOTAL_TEXT, aTotalText[x]);
						_r.addString(Cf.GSM_FIELD_TRANSLATE, aFieldTranslate[x]);
						_result.getItems().add(_r);
					}
				}
			}
			
		} catch (Exception ex) {
			LOG.error(ex.getMessage(), ex);
		}
		return _result;
	}
	
	public ItemRecord dynamicFieldGroup(ItemRecord data, ItemRecord _params) {
		ItemRecord result = new ItemRecord();
		ItemRecord _groupItem = new ItemRecord();
		ItemRecord _totalRecord = new ItemRecord();
		try {
			if (_params == null) {
				return null;
			}
			
			String groupField = _params.getString(Cf.GSM_GROUP_FIELD);
			
			ItemRecord _columns = data.getMapKeyItems().get("_columns");
			if (_columns == null) {
				_columns = ItemRecord.createColumnsItem(data.getItems().get(0));
			}
			int zz = 0;
			int totalzz = 0;
			for (ItemRecord r : data.getItems()) {
				for (ItemRecord _col : _columns.getItems()) {
					if (groupField.equals(_col.getString(Cf.COL_NAME))) {
						String key;
						if (_col.getString(Cf.COL_TYPE).equals(ItemRecord.INTEGER)) {
							key = "" + r.getInteger(groupField);
						} else {
							key = r.getString(groupField);
						}
						
						ItemRecord _gr = _groupItem.getMapKeyItems().get(key);
						if (_gr == null) {
							zz = 1;
							_gr = new ItemRecord();
							_gr.addString(groupField, key);
							_groupItem.getMapKeyItems().put(key, _gr);
						} else {
							zz = 0;
						}
						
						computeTotal(r, _columns, zz, _gr, _params, false);
						totalzz++;
						computeTotal(r, _columns, totalzz, _totalRecord, _params, true);
						
					}
				}
			}
			
			for (Entry<String, ItemRecord> entry : _groupItem.getMapKeyItems().entrySet()) {
				result.getItems().add(entry.getValue());
			}
			
			if (result.getItems().size() > 0) {
				ItemRecord _resColums = ItemRecord.createColumnsItem(result.getItems().get(0));
				result.getMapKeyItems().put("_columns", _resColums);
				if (_totalRecord.getMapItems().size() > 0) {
					result.getMapKeyItems().put("_total", _totalRecord);
				}
			}
			
		} catch (Exception ex) {
			LOG.error(ex.getMessage(), ex);
		}
		return result;
	}
	
	
	private void computeTotal(ItemRecord r, ItemRecord _columns, int record, ItemRecord _total, ItemRecord _params, boolean isTotal) {
		if (! "".equals(_params.getString(Cf.GSM_TOTAL_FIELDS))) {
			
			for (String _values : _params.getString(Cf.GSM_TOTAL_FIELDS).split("[;]")) {
				String[] _elems = _values.split("[:]");
				String _fieldName = oracleString(_elems[0]);
				String _operator = _elems[1];
				
				
				for (ItemRecord _col : _columns.getItems()) {
				
				String colName = _col.getString("NAME");
				String colType = _col.getString("ITEMTYPE");
				
				
				if (colName.equals(_fieldName)) {
					if (_operator.equals(Cf.TOTAL_SUM)) {
						ItemRecord.dynamicSummaryOf(_total, r, colName, colType);
					} else if (_operator.equals(Cf.TOTAL_COUNTER)) {
						colType = ItemRecord.INTEGER;
						ItemRecord.dynamicSummaryOf(_total, 1, colName, colType);
					} else if (_operator.equals(Cf.TOTAL_MAX)) {
						if (colType.equals(ItemRecord.STRING)) {
							if (isTotal) {
								ItemRecord.dynamicMaxf(_total, "", colName, colType);
							} else {
								ItemRecord.dynamicMaxf(_total, r.getString(_fieldName), colName, colType);	
							}
						}
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
	}
	
	private String oracleString(String value) {
		String result = value;
		if (Util.isOracleConnection(_con)) {
			result = value.toUpperCase();
		}
		return result;
	}
	
}
