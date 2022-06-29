package at.goto24.data;


import java.util.Comparator;
import java.util.Date;

import at.goto24.data.ItemRecord;

public class ComparItems implements Comparator<ItemRecord> {

	private String[] orderFields = new String[]{};
	private boolean isDesc = false;

	public ComparItems(String fieldName) {
		this.orderFields = fieldName.split("[,]");
	}
	
	public ComparItems(String fieldName, Boolean isDesc) {
		this.orderFields = fieldName.split("[,]");
		this.isDesc = isDesc;
	}
	
	@Override
	public int compare(ItemRecord o1, ItemRecord o2) {
		int result = 0;
		
		for (String _field : orderFields) {
		
			if (isDesc) {
				if (result == 0) {
					if (o2.getObject(_field) instanceof String) {
						result = o2.getString(_field).compareTo(o1.getString(_field));
					} else if (o2.getObject(_field) instanceof Integer) {
						result = o2.getInteger(_field).compareTo(o1.getInteger(_field));
					} else if (o2.getObject(_field) instanceof Long) {
						result = o2.getLong(_field).compareTo(o1.getLong(_field));
					} else if (o2.getObject(_field) instanceof Double) {
						result = o2.getDouble(_field).compareTo(o1.getDouble(_field));
					} else if (o2.getObject(_field) instanceof Date) {
						result = o2.getDate(_field).compareTo(o1.getDate(_field));
					}
				}
				
			} else {
				if (result == 0) {
					if (o1.getObject(_field) instanceof String) {
						result = o1.getString(_field).compareTo(o2.getString(_field));
					} else if (o1.getObject(_field) instanceof Integer) {
						result = o1.getInteger(_field).compareTo(o2.getInteger(_field));
					} else if (o1.getObject(_field) instanceof Long) {
						result = o1.getLong(_field).compareTo(o2.getLong(_field));
					} else if (o1.getObject(_field) instanceof Double) {
						result = o1.getDouble(_field).compareTo(o2.getDouble(_field));
					} else if (o1.getObject(_field) instanceof Date) {
						result = o1.getDate(_field).compareTo(o2.getDate(_field));
					}
				}
			}
		}
		
		return result;
	}

}
