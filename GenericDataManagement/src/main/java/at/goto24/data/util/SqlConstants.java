package at.goto24.data.util;

import java.sql.Connection;
import java.util.*;
import java.util.regex.Pattern;

import org.apache.log4j.Logger;


public class SqlConstants {
	
	private static final Logger LOG = Logger.getLogger(SqlConstants.class);
	
	public static final String KEY_SEEK = "seek";
	public static final String KEY_REPLACE = "replace";

	private static final String PLACE_HOLDER_TOP_1 = "'OracleTop1'='OracleTop1'";
	private static final String PLACE_HOLDER_TOP_2 = "'OracleTop2'='OracleTop2'";
	private static final String PLACE_HOLDER_TOP_3 = "'OracleTop3'='OracleTop3'";
	private static final String PLACE_HOLDER_TOP_4 = "'OracleTop4'='OracleTop4'";
	private static final String PLACE_HOLDER_TOP_5 = "'OracleTop5'='OracleTop5'";
	private static final String PLACE_HOLDER_TOP_10 = "'OracleTop10'='OracleTop10'";

	/** ReportName, Collection<String'seek'/'replace'><String> */
	private Map<String,Collection<Map<String,String>>> oracleReplacements = new HashMap<>();

	/** ReportName, Collection<String'seek'/'replace'><String> */
	private Map<String,Collection<Map<String,String>>> mySqlReplacements = new HashMap<>();


	public SqlConstants() {
		init();
	}
	
	private void init() {
		
		// TV Scout Report 
		Collection<Map<String,String>> li = new ArrayList<Map<String,String>>();
		Map<String,String> mapReplaceBy = new HashMap<String, String>();
		
	}
	
	/**
	 * Get List by Report Name.  
	 * E.G. design.getName()
	 * 
	 * You get a list of replacements.
	 * @return Map
	 */
	public Collection<Map<String,String>> getOracleReplacements(String reportName) {
		return oracleReplacements.get(reportName);
	}

	/**
	 * Get List by Report Name.
	 * E.G. design.getName()
	 *
	 * You get a list of replacements.
	 * @return Map
	 */
	public Collection<Map<String,String>> getMySqlReplacements(String reportName) {
		return mySqlReplacements.get(reportName);
	}
	
	
	/**
	 * To extend the available replacements.
	 * @param reportName String
	 * @param regExSeekElement seek element using regex formatation
	 * @param replaceElement replace element
	 * @param argList collection of elements
	 * @return collection
	 */
	public Collection<Map<String,String>> addOracleReplacements(String reportName, String regExSeekElement, String replaceElement , Collection<Map<String,String>> argList) {
		Collection<Map<String,String>> li = new ArrayList<Map<String,String>>();
		Map<String,String> mapReplaceBy = new HashMap<String, String>();
		
		if (argList != null) {
			li = argList;
		}
		
		mapReplaceBy = new HashMap<String, String>();
		mapReplaceBy.put(KEY_SEEK, regExSeekElement);
		mapReplaceBy.put(KEY_REPLACE, replaceElement);
		li.add(mapReplaceBy);
		
		oracleReplacements.put(reportName, li);
		
		return li;
	}


	/**
	 * To extend the available replacements.
	 * @param reportName String
	 * @param regExSeekElement seek element using regex formatation
	 * @param replaceElement replace element
	 * @param argList collection of elements
	 * @return collection
	 */
	public Collection<Map<String,String>> addMySqlReplacements(String reportName, String regExSeekElement, String replaceElement , Collection<Map<String,String>> argList) {
		Collection<Map<String,String>> li = new ArrayList<Map<String,String>>();
		Map<String,String> mapReplaceBy = new HashMap<String, String>();

		if (argList != null) {
			li = argList;
		}

		mapReplaceBy = new HashMap<String, String>();
		mapReplaceBy.put(KEY_SEEK, regExSeekElement);
		mapReplaceBy.put(KEY_REPLACE, replaceElement);
		li.add(mapReplaceBy);

		mySqlReplacements.put(reportName, li);

		return li;
	}

	/**
	 * Query parsing: Convert SQL syntax to Oracle or MySql without designated report name (only basic replacements)
	 * @param con
	 * @param queryString
	 * @return
	 */
	public String prepareQuerySyntax(Connection con, String queryString) {
		return prepareQuerySyntax(con, null, queryString);
	}
	
	/**
	 * Query parsing: convert SQL Query syntax to Oracle or MySql Query syntax.
	 * Replacement: Exchange some query elements depending on certain conditions.
	 * @param con Connection
	 * @param reportName  Name of Report or Query
	 * @param queryString Query itself
	 * @return modified query value.
	 */
	public String prepareQuerySyntax(Connection con, String reportName, String queryString) {
		String result = queryString;
		
		if (Util.isOracleConnection(con)) {
			result = prepareOracleQuerySyntax(reportName, result);
		} else if (Util.isMySqlConnection(con)) {
			result = prepareMySqlQuerySyntax(reportName, result);
		}

		return result;
	}
	
	/**
	 * Wrapper method to get connection information.
	 * @param con Connection
	 * @return Boolean
	 */
	public Boolean isOracle(Connection con) {
		return Util.isOracleConnection(con);
	}

	/**
	 * Wrapper method to get connection information.
	 *
	 * @param con Connection
	 * @return Boolean
	 */
	public Boolean isMySql(Connection con) {
		return Util.isMySqlConnection(con);
	}

	private String prepareOracleQuerySyntax(String reportName, String query) {

		if (LOG.isTraceEnabled()) {
			LOG.trace(">>> Oracle:  Query=[" + query +"]");
		}

		query =  query.replaceAll("dbo.", "");

		// 'top 1' was found
		query = query.replaceAll("\\btop 1\\b", "");
		query = query.replaceAll("\\btop 10\\b", "");
		query = query.replaceAll(PLACE_HOLDER_TOP_10, "ROWNUM <= 10");
		query = query.replaceAll(PLACE_HOLDER_TOP_1, "ROWNUM <= 1");

		// concat + ' ' +
		query = query.replaceAll("\\+ '", "|| '");
		query = query.replaceAll("' \\+", "' ||");

		// substring -> substr
		query = query.replaceAll("substring", "substr");

		if (! reportName.isEmpty()) {
			Collection<Map<String, String>> replLi = getOracleReplacements(reportName);
			if (replLi != null) {
				for (Map<String, String> _map : replLi) {
					query = query.replaceAll(_map.get(SqlConstants.KEY_SEEK), _map.get(SqlConstants.KEY_REPLACE));
				}
			}
		}

		if (LOG.isTraceEnabled()) {
			LOG.trace(">>> Oracle: After Parsing Query=[" + query + "]");
		}

		return query;
	}

	private String prepareMySqlQuerySyntax(String reportName, String query) {
		if (LOG.isTraceEnabled()) {
			LOG.trace(">>> MySQL:  Query=[" + query +"]");
		}

		query =  query.replaceAll("dbo.", "");

		// 'top 1' was found
		query = query.replaceAll("\\btop 1\\b", "");
		query = query.replaceAll("\\btop 2\\b", "");
		query = query.replaceAll("\\btop 3\\b", "");
		query = query.replaceAll("\\btop 4\\b", "");
		query = query.replaceAll("\\btop 5\\b", "");
		query = query.replaceAll("\\btop 10\\b", "");

		query = replaceTopClauseForMySql(query, PLACE_HOLDER_TOP_10);
		query = replaceTopClauseForMySql(query, PLACE_HOLDER_TOP_5);
		query = replaceTopClauseForMySql(query, PLACE_HOLDER_TOP_4);
		query = replaceTopClauseForMySql(query, PLACE_HOLDER_TOP_3);
		query = replaceTopClauseForMySql(query, PLACE_HOLDER_TOP_2);
		query = replaceTopClauseForMySql(query, PLACE_HOLDER_TOP_1);
		
		
		// MySql: RB-14865 pre_game_check_report_oracle - bad formatting 
		if ("PGC Report".equals(reportName)) {
			query = query.replaceAll("DBMS_LOB\\.substr\\(stcpgc.COMMENTARY, length\\(commentary\\)\\)", "cast(stcpgc.COMMENTARY as CHAR(200))");
		}

		// concat + ' ' +
		query = query.replaceAll("\\+ '", "|| '");
		query = query.replaceAll("' \\+", "' ||");

		// substring -> substr
		query = query.replaceAll("substring", "substr");

		// as int to as signed
		query = query.replaceAll("as INT\\)", "as SIGNED)");
		query = query.replaceAll("AS INT\\)", "as SIGNED)");

		// oracles "nulls first" to "is null desc"
		query = query.replaceAll("nulls first", "is null desc");

		// remove oracles DBMS_LOB package
		query = query.replaceAll("DBMS_LOB\\.", "");

		// mysql replacement for "charindex" is "locate"
		query = query.replaceAll("charindex", "locate");


		Collection<Map<String, String>> replLi = getMySqlReplacements(reportName);
		if (replLi != null) {
			for (Map<String, String> _map : replLi) {
				query = query.replaceAll(_map.get(SqlConstants.KEY_SEEK), _map.get(SqlConstants.KEY_REPLACE));
			}
		}

		// mysql replacement for "isnull" to "coalesce"
		query = query.replaceAll("isnull\\(", "COALESCE(");
		query = query.replaceAll("ISNULL\\(", "COALESCE(");

		
		
		
		if (LOG.isTraceEnabled()) {
			LOG.trace(">>> MySQL: After Parsing Query=[" + query + "]");
		}
		return query;
	}

	private static String replaceTopClauseForMySql(String query, String toReplace) {

		// Usages of clause:
		// a) where 'OracleTop1'='OracleTop1'
		// b) WHERE 'OracleTop1'='OracleTop1'
		// c) and 'OracleTop1'='OracleTop1'

		// We need to check the 'previous' word before the replacement marker and remove it
		// -> We will do this only in case of keywords 'where' and 'and'

		String[] tokens = query.split(Pattern.quote(toReplace));
		int size = tokens.length;
		List<String> resultTokens = new ArrayList<>();

		String limitToAdd = "";
		switch (toReplace)
		{
			case PLACE_HOLDER_TOP_10:
				limitToAdd = "LIMIT 10";
				break;
			case PLACE_HOLDER_TOP_5:
				limitToAdd = "LIMIT 5";
				break;
			case PLACE_HOLDER_TOP_4:
				limitToAdd = "LIMIT 4";
				break;
			case PLACE_HOLDER_TOP_3:
				limitToAdd = "LIMIT 3";
				break;
			case PLACE_HOLDER_TOP_2:
				limitToAdd = "LIMIT 2";
				break;
			default:
				limitToAdd = "LIMIT 1";
				break;
		}

		for (int i = 0; i < size; i++) {
			String token = tokens[i];

			// Check if token 'ends' with one of the relevant keywords and remove it
			String lower = token.toLowerCase();
			String trimmed = lower.trim();
			if (trimmed.endsWith("where")) {
				int idx = lower.lastIndexOf("where");
				token = token.substring(0, idx) + token.substring(idx+5);
			} else if (trimmed.endsWith("and")) {
				int idx = lower.lastIndexOf("and");
				token = token.substring(0, idx) + token.substring(idx+3);
			}

			resultTokens.add(token);

			if (i < size-1) {
				resultTokens.add(limitToAdd);
			}
		}

		query = String.join("", resultTokens);

		return query;
	}



}
