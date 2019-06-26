package com.mysql;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import com.mysql.jdbc.PreparedStatement;
import com.util.ExcelUtil;

/**
 * 
 * Excel导入Mysql
 * 
 * @version 1.0
 * @author wangcy
 * @date 2019年6月26日 下午4:24:07
 */
public class ExcelToMysql {

	
	static final String DB_URL = "jdbc:mysql://localhost:3306/test";	// 数据库 URL
	static final String USER = "root";						// 数据库的用户名
	static final String PASS = "123456";					// 数据库的密码
	private static final String PATH = "d:\\2019日历.xlsx";	// Excel文件所在的路径
	private static final String TABLE = "calendar";			// 数据库表名
	private static final String FIELD_01 = "date";			// 导入字段
	private static final String FIELD_02 = "type";
	
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
			StringBuffer sql = new StringBuffer("insert into ");
			sql.append(TABLE);
			sql.append("(");
			for (int i = 0; i < FIELDLIST.size(); i++) {
				sql.append(FIELDLIST.get(i) + ",");
			}
			sql.deleteCharAt(sql.length()-1);
			sql.append(") values ");
			// 拼接数据
			List<Map<String, Object>> dataList = ExcelUtil.getExcelData(PATH);
			if (dataList!=null && dataList.size()!=0) {
				for (Map<String, Object> map : dataList) {
					sql.append("(");
					for (int i = 0; i < FIELDLIST.size(); i++) {
						sql.append("'" + map.get(FIELDLIST.get(i)) + "',");
					}
					sql.deleteCharAt(sql.length()-1);
					sql.append("),");
				}
			}
			sql.deleteCharAt(sql.length()-1);
			pstmt = (PreparedStatement) connection.prepareStatement(sql.toString());
			pstmt.execute();
			System.out.println("导入成功");
		} catch (SQLException e) {
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
