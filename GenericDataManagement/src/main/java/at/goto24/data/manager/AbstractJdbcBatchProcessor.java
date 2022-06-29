package at.goto24.data.manager;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedList;
import java.util.List;

import at.goto24.data.util.Util;

public abstract class AbstractJdbcBatchProcessor <T extends Number, R>
{
	private int listSize, batchSize, startIndex, maxToIndex;
	private List<T> idList;
	private String query;
	private String idListPlaceholderName;

    public void init(List<T> idList, String query, int batchSize,
    		String idListPlaceholderName) {
        this.idList = idList;
        this.listSize = idList == null ? 0 : idList.size();
        
        this.query = query;
        
        this.batchSize = batchSize;
        this.maxToIndex = batchSize;
        
        this.idListPlaceholderName = idListPlaceholderName;
    }
    
    public void init(Collection<T> idList, String query, int batchSize,
    		String idListPlaceholderName) {
        this.idList = new ArrayList<>(idList);
        this.listSize = idList == null ? 0 : idList.size();

        this.query = query;

        this.batchSize = batchSize;
        this.maxToIndex = batchSize;
        
        this.idListPlaceholderName = idListPlaceholderName;
    }
    
    public List<R> processQuery(Connection dataSource) throws SQLException {
        List<R> result = new LinkedList<>();
        
        if (idList != null && ! idList.isEmpty()) {
            while (true) {

                if (maxToIndex > listSize) {
                    maxToIndex = listSize;
                }

                // Split id list into sub list with size 'batchSize'           
                List<T> subIdList = idList.subList(startIndex, maxToIndex);
                
                String adaptedQuery = query.replace(idListPlaceholderName, Util.getPlaceHolders(subIdList.size()));
                
                try (PreparedStatement ps = dataSource.prepareStatement(adaptedQuery))
                {
                	for (int i=0;i<subIdList.size();i++)
                	{
                		ps.setObject(i+1, subIdList.get(i));
                	}
                	
                	try (ResultSet rs = ps.executeQuery())
                	{
                		while (rs.next())
                		{
                			R constructedItem = constructItem (rs);
                			result.add(constructedItem);
                		}
                	}
                }

                startIndex = maxToIndex;
                maxToIndex += batchSize;

                // Continue with querying as long 'listSize' is not yet reached            
                if (startIndex >= listSize) {
                    break;
                }
            }
        }
        
        return result;
    }
    
    public abstract R constructItem(ResultSet rs) throws SQLException;
}