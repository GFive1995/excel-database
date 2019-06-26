package com.mysql;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mysql.jdbc.PreparedStatement;

/**
 * 
 * Mysql导入Excel
 * 
 * @version 1.0
 * @author wangcy
 * @date 2019年6月26日 下午4:24:23
 */
public class MysqlToExcel {
	// 数据库 URL
    static final String DB_URL = "jdbc:mysql://localhost:3306/test";
    // 数据库的用户名与密码，需要根据自己的设置
    static final String USER = "root";
    static final String PASS = "123456";
    // Excel文件所在的路径
    private static String PATH = "d:\\2019日历.xls";	
	
	// 数据库表
	private static final String TABLE = "calendar";
	// 数据库字段
	private static String FIELD_01 = "date";
	private static String FIELD_02 = "type";
	
	private static List<String> FIELDS = new ArrayList<>();
	static {
		FIELDS.add(FIELD_01);
		FIELDS.add(FIELD_02);
	}
    
    public static void main(String[] args) {
    	Connection connection = null;
    	PreparedStatement pstmt = null;
    	FileOutputStream outputStream = null;
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
				for (int i = 0; i < FIELDS.size(); i++) {
					paraMap.put(FIELDS.get(i), rs.getString(FIELDS.get(i)));
				}
				dataList.add(paraMap);
			}
            rs.close();
			// 创建HSSFWorkbook对象
			Workbook workbook = null;
			if (PATH.endsWith(".xls")) {
				workbook = new HSSFWorkbook();
			} else if (PATH.endsWith(".xlsx")) {
				workbook = new XSSFWorkbook();
			}
			// 创建HSSFSheet对象
			HSSFSheet sheet = (HSSFSheet) workbook.createSheet("sheet");
			// Excel表头
			HSSFRow row0 = sheet.createRow(0);
			for (int i = 0; i < FIELDS.size(); i++) {
				HSSFCell cell0 = row0.createCell(i);
				cell0.setCellValue(FIELDS.get(i));
			}
			// Excel数据
			for (int i = 0; i < dataList.size(); i++) {
				HSSFRow rows = sheet.createRow(i+1);
				for (int j = 0; j < FIELDS.size(); j++) {
					HSSFCell cells = rows.createCell(j);
					cells.setCellValue(dataList.get(i).get(FIELDS.get(j)).toString());
				}
			}
			// 输出文件
			outputStream = new FileOutputStream(PATH);
			workbook.write(outputStream);
			outputStream.flush();
			System.out.println("导出成功");
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}  catch (Exception e) {
			e.printStackTrace();
		}
	}
    
}
