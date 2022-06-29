package at.goto24.data.storage;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;

import org.apache.log4j.Logger;

import at.goto24.data.ItemRecord;



/**
 * It makes possible to store or to read serialized ItemRecord Object to the or from the file system!
 * @author j.toescher
 * @version 1.0.0
 * @since 01.06.2017
 * 
 * Sonar fixed 31.01.2017
 */
public class ItemStorage {
	
	private static final Logger LOG = Logger.getLogger(ItemStorage.class);
	
	public static boolean save(ItemRecord item, String fileName) {
		boolean ok = false;
		
		try (ObjectOutputStream out = 
				new ObjectOutputStream(new BufferedOutputStream(
						new FileOutputStream(fileName)))) {
		out.writeObject(item);
		ok = true;
		} catch (Exception ex) {
			ok = false;
			LOG.error(ex.getMessage(), ex);
		}
		
		return ok;
	}
	
	
	public static boolean saveObject(Object obj, String fileName) {
		boolean ok = false;
		
		try (ObjectOutputStream out = 	new ObjectOutputStream(new BufferedOutputStream(
						new FileOutputStream(fileName)))) {
		out.writeObject(obj);
		ok = true;
		} catch (Exception ex) {
			ok = false;
			LOG.error(ex.getMessage(), ex);
		}
		
		return ok;
	}
	
	
	
	public static ItemRecord read(String fileName) {
		ItemRecord result = null;
		try (ObjectInputStream 	in = new ObjectInputStream(
				new BufferedInputStream(
						new FileInputStream(fileName)))) {
		Object object = in.readObject();
		
		if (object != null) {
			result = (ItemRecord) object;
		}
		
		} catch (Exception ex) {
			LOG.error(ex.getMessage(), ex);
		}
		return result;
	}

	
	public static Object readObject(String fileName) {
		Object result = null;
		try (ObjectInputStream	in = new ObjectInputStream(
				new BufferedInputStream(
						new FileInputStream(fileName)))) {
		Object object = in.readObject();
		
		if (object != null) {
			result = object;
		}
		
		} catch (Exception ex) {
			LOG.error(ex.getMessage(), ex);
		}
		return result;
	}
	
	
}
