package com.mysql;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mysql.jdbc.PreparedStatement;

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
	private static final String PATH = "d:\\2019日历.xls";	// Excel文件所在的路径
	private static final String TABLE = "calendar";			// 数据库表名
	private static final String FIELD_01 = "date";			// 数据库字段
	private static final String FIELD_02 = "type";
	
	private static List<String> FIELDLIST = new ArrayList<>();
	static {
		FIELDLIST.add(FIELD_01);
		FIELDLIST.add(FIELD_02);
	}
	
	public static void main(String[] args) {
		Connection connection = null;
		PreparedStatement pstmt = null;
		Workbook workbook = null;
		try {
			// 输入文件
			FileInputStream inputStream = new FileInputStream(PATH);
			if (PATH.endsWith(".xls")) {
				workbook = new HSSFWorkbook(inputStream);
			} else if (PATH.endsWith(".xlsx")) {
				workbook = new XSSFWorkbook(inputStream);
			}
			// 打开链接
			System.out.println("连接数据库...");
			connection = DriverManager.getConnection(DB_URL, USER, PASS);
			List<Map<String, Object>> dataList = new ArrayList<>();
			List<String> fieldList = new ArrayList<String>();
			StringBuffer sql = new StringBuffer("insert into ");
			sql.append(TABLE);
			sql.append("(");
			for (int i = 0; i < FIELDLIST.size(); i++) {
				sql.append(FIELDLIST.get(i) + ",");
			}
			sql.deleteCharAt(sql.length()-1);
			sql.append(") values ");
			// 获取Excel文档中第一个表单
			Sheet sheet = workbook.getSheetAt(0);
			// 获取Excel第一行名称
			Row row0 = sheet.getRow(0);
			for (Cell cell : row0) {
				fieldList.add(cell.toString());
			}
			int rows = sheet.getLastRowNum() + 1;
			int cells = fieldList.size();
			for (int i = 1; i < rows; i++) {
				Row row = sheet.getRow(i);
				Map<String, Object> paraMap = new HashMap<>();
				for (int j = 0; j < cells; j++) {
					Cell cell = row.getCell(j);
					if (cell != null && !cell.equals("")) {
						paraMap.put(fieldList.get(j), cell.toString());
					}
				}
				dataList.add(paraMap);
			}
			for (Map<String, Object> map : dataList) {
				sql.append("(");
				for (int i = 0; i < FIELDLIST.size(); i++) {
					sql.append("'" + map.get(FIELDLIST.get(i)) + "',");
				}
				sql.deleteCharAt(sql.length()-1);
				sql.append("),");
			}
			sql.deleteCharAt(sql.length()-1);
			pstmt = (PreparedStatement) connection.prepareStatement(sql.toString());
			pstmt.execute();
			System.out.println("导入成功");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (workbook != null) {
				try {
					workbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
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
