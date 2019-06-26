package com.mysql;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.mysql.jdbc.PreparedStatement;
import com.util.ExcelUtil;

/**
 * 
 * Mysql导入Excel
 * 
 * @version 1.0
 * @author wangcy
 * @date 2019年6月26日 下午4:24:23
 */
public class MysqlToExcel { 
	
    static final String DB_URL = "jdbc:mysql://localhost:3306/test";	// 数据库 URL
    static final String USER = "root";					 	// 数据库的用户名
    static final String PASS = "123456";				 	// 数据库的密码
    private static String PATH = "d:\\2019日历.xlsx";			// Excel文件所在的路径
    private static String FILE_NAME = "2019日历.xlsx";		// 文件名
	private static final String TABLE = "calendar";			// 数据库表名
	private static String FIELD_01 = "date";				// 数据库字段
	private static String FIELD_02 = "type";
	
	private static List<String> FIELDLIST = new ArrayList<>();
	static {
		FIELDLIST.add(FIELD_01);
		FIELDLIST.add(FIELD_02);
	}
    
    public static void main(String[] args) {
    	Connection connection = null;
    	PreparedStatement pstmt = null;
    	try {
    		// 打开链接
    		System.out.println("连接数据库...");
			connection = DriverManager.getConnection(DB_URL, USER, PASS);
			String sql = "SELECT * from " + TABLE;
			pstmt = (PreparedStatement) connection.prepareStatement(sql);
            ResultSet rs = pstmt.executeQuery(sql);
            List<Map<String, Object>> dataList = new ArrayList<>();
			while (rs.next()) {
				Map<String, Object> paraMap = new HashMap<String, Object>();
				for (int i = 0; i < FIELDLIST.size(); i++) {
					paraMap.put(FIELDLIST.get(i), rs.getString(FIELDLIST.get(i)));
				}
				dataList.add(paraMap);
			}
            rs.close();
            ExcelUtil.downloadExcel(FILE_NAME, PATH, FIELDLIST, FIELDLIST, dataList);
			System.out.println("导出成功");
		} catch (SQLException e) {
			e.printStackTrace();
		}  catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (connection != null) {
				try {
					connection.close();
				} catch (SQLException e) {
					e.printStackTrace();
				}
			}
			if (pstmt != null) {
				try {
					pstmt.close();
				} catch (SQLException e) {
					e.printStackTrace();
				}
			}
		}
	}
    
}
