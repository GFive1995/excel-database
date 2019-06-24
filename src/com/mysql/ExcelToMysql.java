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

public class ExcelToMysql {

	// 数据库 URL
	static final String DB_URL = "jdbc:mysql://localhost:3306/test";
	// 数据库的用户名与密码，需要根据自己的设置
	static final String USER = "root";
	static final String PASS = "123456";
	// Excel文件所在的路径
	private static String PATH = "d:\\2019日历.xls";

	// 数据库字段
	private static String FIELD_01 = "date";
	private static String FIELD_02 = "type";
	
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
			StringBuffer sql = new StringBuffer("insert into calendar("+FIELD_01+", "+FIELD_02+") values ");
			// 获取Excel文档中第一个表单
			Sheet sheet = workbook.getSheetAt(0);
			Row row0 = sheet.getRow(0);
			for (Cell cell : row0) {
				fieldList.add(cell.toString());
			}
			sql.substring(0, sql.length() - 1);
			System.out.println(sql);
			System.out.println(fieldList);
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
				sql.append("('" + map.get(FIELD_01) + "',");
				sql.append("" + map.get(FIELD_02) + "),");
			}
			String exeSql = sql.substring(0, sql.toString().length() - 1);
			pstmt = (PreparedStatement) connection.prepareStatement(exeSql);
			pstmt.execute();
			System.out.println("导入成功");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
