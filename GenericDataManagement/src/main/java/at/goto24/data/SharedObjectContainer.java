package at.goto24.data;

/**
 * Singleton Class to manage meta central data to share cross over report instances.
 * Available up version 4.0.68
 * 
 * @author j.toescher
 * @version 1.0.0
 * 
 */

import java.util.Date;

import org.apache.log4j.Logger;

public class SharedObjectContainer {
	private static final Logger LOG = Logger.getLogger(SharedObjectContainer.class);
	
	private static SharedObjectContainer me = null;
	private ItemRecord item = new ItemRecord();
	private Date createDate = null;
	
	private SharedObjectContainer() {
		item = new ItemRecord();
	}
	
	public synchronized static SharedObjectContainer getInstance() {
		Date toDay = new Date();
		if (me == null) {
			me = new SharedObjectContainer();
			me.setCreateDate(toDay);
		} else {
			Date createDate = me.getCreateDate();
			Long dateDiff = toDay.getTime() - createDate.getTime();
			if (dateDiff > 0) {
				long hour = 1000L * 60L * 60L;
				long x24Hours = 24 * hour;
				
				if (dateDiff > x24Hours) {
					me = new SharedObjectContainer();
					me.setCreateDate(toDay);
				}
			}
			
		}
		return me;
	}
	
	public synchronized boolean containsDataByName(String key) {
		return item.getMapKeyItems().containsKey(key);
	}
	
	public synchronized ItemRecord createItemContainerByName(String key) {
		ItemRecord retItem = item.getMapKeyItems().get(key);
		if (retItem == null) {
			retItem = new ItemRecord();
			item.getMapKeyItems().put(key, retItem);
		}
		return retItem;
	}
	
	public synchronized ItemRecord storeItemContainerByName(String key, ItemRecord dataItem) {
		item.getMapKeyItems().put(key, dataItem);
		return item.getMapKeyItems().get(key);
	}
	
	
	public synchronized ItemRecord getItemContainerByName(String key) {
		ItemRecord retItem = item.getMapKeyItems().get(key);
		return retItem;
	}
	
	
	public synchronized ItemRecord getBaseItem() {
		return item;
	}
	
	public synchronized boolean clearAll() {
		boolean back = false;
		try {
			item.getMapKeyItems().clear();
			item.getMapItems().clear();
			
			item = new ItemRecord();
			
		} catch (Exception ex) {
			LOG.error(ex.getMessage(), ex);
		}
		return back;
	}

	public synchronized Date getCreateDate() {
		return createDate;
	}

	public synchronized void setCreateDate(Date createDate) {
		this.createDate = createDate;
	}
	
}
