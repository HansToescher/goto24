package at.goto24.data.util;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

import at.goto24.data.ItemRecord;

/**
 * Utility Class - reading basic data like groups, sports etc...
 * Reading Base Data crossing over the all reports
 * @author j.toescher
 * @version 1.0.0
 * @since 10.03.2017
 */
public class ReadBaseData {
	private Connection _con = null;
	private SqlConstants sqlConstants;
	
	public static String F_ID = "F_ID";
	public static String F_NAME = "F_NAME";
	
	public ReadBaseData(Connection dataSource, SqlConstants sqlConstants) {
		super();
		_con = dataSource;
		this.sqlConstants = sqlConstants;
	}
	
	
	/**
	 * Read Group Name by Id
	 * F_ID, F_NAME
	 * @param groupId
	 * @return ItemRecord
	 * @throws SQLException
	 */
	public ItemRecord readGroupById(Integer groupId) throws SQLException {
		ItemRecord data = new ItemRecord();

		String sQuery = "select ug.IDUserGroup, ug.description "
				+"from UserGroup ug "
				+"where ug.IDUserGroup = ? "; // groupId 
		
		int sx = 1;
		
		try (PreparedStatement stmt = _con.prepareStatement(sqlConstants.prepareQuerySyntax(_con, sQuery))) {
		stmt.setInt(sx++, groupId);
		
		int rx = 1;
			
		try (ResultSet rs =  stmt.executeQuery()) {
			if (rs != null) {
				while (rs.next()) {
					rx = 1;
					ItemRecord r = new ItemRecord();
					r.addItem(F_ID, rs.getInt(rx++), ItemRecord.INTEGER);
					r.addItem(F_NAME, rs.getString(rx++), ItemRecord.STRING);
					
					data = r;
				}
			}
		}
			
			
		
		}
		return data;
	}
}