package at.goto24.data;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;


/**
 * General Data Container of some Java Reports.
 * @author j.toescher
 * @version 1.0.0
 */
public class ItemRecord implements Serializable {
	private static final long serialVersionUID = -6085647849565794218L;
	public static final String BOOLEAN = "Boolean";
	public static final String LONG = "Long";
	public static final String DOUBLE = "Double";
	public static final String DATE2 = "Date";
	public static final String INTEGER = "Integer";
	public static final String STRING = "String";
	public static final String SHORT = "Short";
	public static final String DYNAMIC_COUNTER = "_counter";
	
	private Integer id;
	private Double doubleValue;
	private Integer intValue;
	private Short shortValue;
	private Boolean check;
	private Date date = null;
	private Long longValue;
	private String text;
	private String name;
	private String className;
	private transient Object tag;
	private transient Object obj;
	
	private Map<String, ItemRecord> mapItems = new HashMap<>();
	private Map<String, ItemRecord> mapKeyItems = new HashMap<>();
	private List<ItemRecord> items = new ArrayList<>();
	
	public ItemRecord() {
		super();
	}
	
	 
	public String toString() {
		StringBuilder sb = new StringBuilder();
		
		sb.append(this.name == null ? "" : this.name);
        if (this.tag != null) {
            sb.append("=");
            sb.append(this.tag);
        } else {
            sb.append("Item");
        }
        
		if (this.mapItems.size() > 0) {
			sb.append(" [");
			int _x = 0;
			for (Entry<String, ItemRecord> _entry : mapItems.entrySet()) {
				if (_x > 0) sb.append(", ");
				sb.append(_entry.getValue().toString());
				_x++;
			}
			sb.append("]; ");
		}
		
		if (this.mapKeyItems.size() > 0) {
			sb.append(" {");
			int x = 0;
			for (Entry<String, ItemRecord> _entry : mapKeyItems.entrySet()) {
				if (x > 0) sb.append("; ");
				sb.append(" key:[" + _entry.getKey() + "]=");
				sb.append(_entry.getValue().toString());
				x++;
			}
			
			sb.append(" }");
		}
		
		return sb.toString();
	}
	
	
	public Integer getId() {
		return id;
	}
	public void setId(Integer id) {
		this.id = id;
	}
	public Double getDoubleValue() {
		return doubleValue;
	}
	public void setDoubleValue(Double doubleValue) {
		this.doubleValue = doubleValue;
	}
	public Integer getIntValue() {
		return intValue;
	}
	public void setIntValue(Integer intValue) {
		this.intValue = intValue;
	}
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getClassName() {
		return className;
	}
	public void setClassName(String className) {
		this.className = className;
	}
	public Object getTag() {
		return tag;
	}
	public void setTag(Object tag) {
		this.tag = tag;
	}
	public Map<String, ItemRecord> getMapItems() {
		return mapItems;
	}
	public void setMapItems(Map<String, ItemRecord> mapItems) {
		this.mapItems = mapItems;
	}
	public List<ItemRecord> getItems() {
		return items;
	}
	public void setItems(List<ItemRecord> items) {
		this.items = items;
	}
	public String getText() {
		return text;
	}
	public void setText(String text) {
		this.text = text;
	}

	public Date getDate() {
		return date;
	}

	public void setDate(Date date) {
		this.date = date;
	}
	
	public ItemRecord addDate(String key, Date arg) {
		return addItem(key, arg, ItemRecord.DATE2);
	}
	
	public ItemRecord addDouble(String key, Double arg) {
		return addItem(key, arg, ItemRecord.DOUBLE);
	}
	
	public ItemRecord addString(String key, String arg) {
		return addItem(key, arg, ItemRecord.STRING);
	}
	
	public ItemRecord addLong(String key, Long arg) {
		return addItem(key, arg, ItemRecord.LONG);
	}
	
	public ItemRecord addInteger(String key, Integer arg) {
		return addItem(key, arg, ItemRecord.INTEGER);
	}
	
	public ItemRecord addItem(String key, Object o, String type) {
		ItemRecord _rec =  mapItems.get(key);
		if (_rec == null) {
			_rec = new ItemRecord();
			mapItems.put(key, _rec);
		}
		_rec.setClassName(type);
		_rec.setName(key);
		_rec.setTag(o);
		if (STRING.equals(type)) {
			_rec.setText((String) o);
		} else if (INTEGER.equals(type)) {
			_rec.setIntValue((Integer) o) ;
		} else if (SHORT.equals(type)) {
			_rec.setShortValue((Short) o) ;
		} else if (DATE2.equals(type)) {
			_rec.setDate((Date)o);
		} else if (DOUBLE.equals(type)) {
			_rec.setDoubleValue((Double) o);
		} else if (LONG.equals(type)) {
			_rec.setLongValue((Long) o);
		} else if (BOOLEAN.equals(type)) {
			_rec.setCheck((Boolean) o);
		} else {
			if (type != null) {
				_rec.setClassName(type);
			}
		}
		return _rec;
	}

	
	public ItemRecord addKeyItem(String key) {
		ItemRecord _rec = mapKeyItems.get(key);
		if (_rec == null) {
			_rec = new ItemRecord();
			mapKeyItems.put(key, _rec);
		}
		return _rec;
	}
	
	/**
	 * To add new Element as Record Field
	 * @param key String as Name of the Object
	 * @param o Object as value to be managed
	 * @param type String e.g. String, Integer, Double ...
	 * @param isSave Boolean 'true' means, the first written value can not be overwritten by other values.
	 */
	public void addItemSave(String key, Object o, String type, Boolean isSave) {
		
		if (o == null) {
			return;
		}
		
		ItemRecord _rec =  mapItems.get(key);
		if (_rec == null) {
			_rec = new ItemRecord();
			mapItems.put(key, _rec);
		} else {
			if (isSave) {
				if (_rec.getTag() != null) {
					return;
				}
			}
		}
		_rec.setClassName(type);
		_rec.setName(key);
		_rec.setTag(o);
		if (STRING.equals(type)) {
			_rec.setText((String) o);
		} else if (INTEGER.equals(type)) {
			_rec.setIntValue((Integer) o) ;
		} else if (SHORT.equals(type)) {
			_rec.setShortValue((Short) o) ;
		} else if (DATE2.equals(type)) {
			_rec.setDate((Date)o);
		} else if (DOUBLE.equals(type)) {
			_rec.setDoubleValue((Double) o);
		} else if (LONG.equals(type)) {
			_rec.setLongValue((Long) o);
		} else if (BOOLEAN.equals(type)) {
			_rec.setCheck((Boolean) o);
		} else {
			if (type != null) {
				_rec.setClassName(type);
			}
		}
	}
	
	public Integer getInteger(String key) {
		ItemRecord rec = mapItems.get(key);
		
		if (rec != null) {
			return rec.getIntValue();
		}
		return 0;
	}
	
	public Short getShort(String key) {
		ItemRecord rec = mapItems.get(key);
		
		if (rec != null) {
			return rec.getShortValue();
		}
		return 0;
	}
	
	
	public String getString(String key) {
		ItemRecord rec = mapItems.get(key);
		
		if (rec != null) {
			return rec.getText();
		}
		return "";
	}
	
	public Date getDate(String key) {
		ItemRecord rec = mapItems.get(key);
		
		if (rec != null) {
			return rec.getDate();
		}
		return null;
	}
	
	public Double getDouble(String key) {
		ItemRecord rec = mapItems.get(key);
		
		if (rec != null) {
			return rec.getDoubleValue();
		}
		return 0D;
	}
	
	public Boolean getBoolean(String key) {
		ItemRecord rec = mapItems.get(key);
		
		if (rec != null) {
			return rec.getCheck();
		}
		return null;
	}
	
	
	public Long getLong(String key) {
		ItemRecord rec = mapItems.get(key);
		
		if (rec != null) {
			return rec.getLongValue();
		}
		return 0L;
	}
	
	
	public Object getObject(String key) {
		ItemRecord rec = mapItems.get(key);
		
		if (rec != null) {
			return rec.getTag();
		}
		return null;
	}

	public Boolean getCheck() {
		return check;
	}

	public void setCheck(Boolean check) {
		this.check = check;
	}

	public Long getLongValue() {
		return longValue;
	}

	public void setLongValue(Long longValue) {
		this.longValue = longValue;
	}

	public Short getShortValue() {
		return shortValue;
	}

	public void setShortValue(Short shortValue) {
		this.shortValue = shortValue;
	}
	
	/**
	 * Summary Aggregate Method
	 * @param sr
	 * @param r
	 * @param fieldName
	 * @param type
	 */
	public static void dynamicSummaryOf(ItemRecord sr, ItemRecord r, String fieldName, String type) {
		
		if (type.equals(ItemRecord.DOUBLE)) {
			Double sum = sr.getDouble(fieldName).doubleValue() + r.getDouble(fieldName).doubleValue();
			sr.addItem(fieldName, sum, ItemRecord.DOUBLE);
		}
		
		if (type.equals(ItemRecord.INTEGER)) {
			Integer sum = sr.getInteger(fieldName).intValue() + r.getInteger(fieldName).intValue();
			sr.addItem(fieldName, sum, ItemRecord.INTEGER);
		}

		if (type.equals(ItemRecord.LONG)) {
			Long sum = sr.getLong(fieldName).longValue() + r.getLong(fieldName).longValue();
			sr.addItem(fieldName, sum, ItemRecord.LONG);
		}
		
		// specially function to count selected records.
		if (DYNAMIC_COUNTER.equals(type)) {
			Integer sum = sr.getInteger(fieldName).intValue() + 1;
			sr.addItem(fieldName, sum, ItemRecord.INTEGER);
		}
		
	}
	
	/**
	 * Summary Aggregate Method
	 * @param sr
	 * @param oValue
	 * @param fieldName
	 * @param type
	 */
	public static void dynamicSummaryOf(ItemRecord sr, Object oValue, String fieldName, String type) {
		if (type.equals(ItemRecord.DOUBLE)) {
			Double sum = sr.getDouble(fieldName).doubleValue() + ((Double) oValue).doubleValue();
			sr.addItem(fieldName, sum, ItemRecord.DOUBLE);
		}
		
		if (type.equals(ItemRecord.INTEGER)) {
			Integer sum = sr.getInteger(fieldName).intValue() + ((Integer) oValue).intValue();
			sr.addItem(fieldName, sum, ItemRecord.INTEGER);
		}
		
		if (type.equals(ItemRecord.LONG)) {
			Long sum = sr.getLong(fieldName).longValue() + ((Long) oValue).longValue();
			sr.addItem(fieldName, sum, ItemRecord.LONG);
		}
		
		// specially function to count selected records.
		if (DYNAMIC_COUNTER.equals(type)) {
			Integer sum = sr.getInteger(fieldName).intValue() + 1;
			sr.addItem(fieldName, sum, ItemRecord.INTEGER);
		}
	}
	
	public static void dynamicMaxf(ItemRecord sr, Object oValue, String fieldName, String type) {
		if (type.equals(ItemRecord.DOUBLE)) {
			Double sum = sr.getDouble(fieldName).doubleValue();
			if (sum < ((Double) oValue).doubleValue()) {
				sr.addItem(fieldName, ((Double) oValue).doubleValue(), ItemRecord.DOUBLE);
			}
		}
		
		if (type.equals(ItemRecord.INTEGER)) {
			Integer sum = sr.getInteger(fieldName).intValue(); 
			if (sum < ((Integer) oValue).intValue()) {		
				sr.addItem(fieldName, ((Integer) oValue).intValue(), ItemRecord.INTEGER);
			}
		}
		
		if (type.equals(ItemRecord.LONG)) {
			Long sum = sr.getLong(fieldName).longValue();
			if (sum < ((Long) oValue).longValue()) {
				sr.addItem(fieldName, ((Long) oValue).longValue(), ItemRecord.LONG);
			}
		}
		
		if (type.equals(ItemRecord.STRING)) {
			String sum = sr.getString(fieldName).trim();
			if ("".equals(sum)) {
				sr.addString(fieldName, "" + oValue);
			} else {
				if (sum.length() < (""+oValue).length()) {
					sr.addString(fieldName, "" + oValue);
				}
			}
		}
		
		// specially function to count selected records.
		if (DYNAMIC_COUNTER.equals(type)) {
			Integer sum = sr.getInteger(fieldName).intValue() + 1;
			sr.addItem(fieldName, sum, ItemRecord.INTEGER);
		}
	}
	
	/**
	 * Average Aggregate Method
	 * @param sr
	 * @param oValue
	 * @param fieldName
	 * @param type (only for Double,Integer,Long
	 */
	public static void averageOf(ItemRecord sr, Object oValue, String fieldName, String type) {
		
		if (type.equals(ItemRecord.DOUBLE)) {
			Double sum = (sr.getDouble(fieldName).doubleValue() + ((Double) oValue).doubleValue()) / 2;
			sr.addItem(fieldName, sum, ItemRecord.DOUBLE);
		}
		
		if (type.equals(ItemRecord.INTEGER)) {
			Integer sum = (sr.getInteger(fieldName).intValue() + ((Integer) oValue).intValue()) / 2 ;
			sr.addItem(fieldName, sum, ItemRecord.INTEGER);
		}
		
		if (type.equals(ItemRecord.LONG)) {
			Long sum = (sr.getLong(fieldName).longValue() + ((Long) oValue).longValue()) / 2L;
			sr.addItem(fieldName, sum, ItemRecord.LONG);
		}
		
	}


	public Map<String, ItemRecord> getMapKeyItems() {
		return mapKeyItems;
	}


	public void setMapKeyItems(Map<String, ItemRecord> mapKeyItems) {
		this.mapKeyItems = mapKeyItems;
	}
	
	/**
	 * Creates a field descriptor for GenericReport to build result Excel Sheets and more ..
	 * Please place the result in Itemrecord.getMapKeyItems().put("_columns",returnValue)...
	 * 
	 * @param r ItemRecord as data record
	 * @return ItemRecord as Field Descriptor
	 */
    public static ItemRecord createColumnsItem(ItemRecord r) {
        ItemRecord columns = new ItemRecord();
        
        int x = 0;
        for (Entry<String,ItemRecord> _entry : r.getMapItems().entrySet()) {
            x++;
            ItemRecord t = _entry.getValue();
            String _name = t.getName();
            String _itemType = t.getClassName();
            
            ItemRecord c = new ItemRecord();
            c.addItem("NAME", _name, ItemRecord.STRING);
            c.addItem("ITEMTYPE", _itemType, ItemRecord.STRING);
            c.addItem("TYPE", _itemType, ItemRecord.STRING);
            c.addItem("INDEX", x, ItemRecord.INTEGER);
            
            columns.getItems().add(c);
            
        }
        
        return columns;
    }


	public Object getObj() {
		return obj;
	}


	public void setObj(Object obj) {
		this.obj = obj;
	}
    
    
}