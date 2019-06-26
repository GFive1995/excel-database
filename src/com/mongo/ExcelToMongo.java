package com.mongo;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.bson.Document;

import com.mongodb.MongoClient;
import com.mongodb.MongoCredential;
import com.mongodb.ServerAddress;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;

/**
 * 
 * Excel数据导入MongoDB
 * 
 * @version 1.0
 * @author wangcy
 * @date 2019年6月26日 下午5:41:38
 */
public class ExcelToMongo {

	private static Integer PORT = 27017; // 端口号
	private static String IP = "localhost"; // Ip
	private static String DATABASE = "database"; // 数据库名称
	private static String USERNAME = "username"; // 用户名
	private static String PASSWORD = "password"; // 密码
	private static String COLLECTION = "calendar"; // 文档名称
	private static String ADDRESS = "d:\\2019日历.xls"; // Excel文件所在的路径

	public static void main(String[] args) {
		Workbook workbook = null;
		MongoClient mongoClient = null;
		try {
			// 根据输入流导入Excel产生Workbook对象
			FileInputStream inputStream = new FileInputStream(ADDRESS);
			if (ADDRESS.endsWith(".xls")) {
				workbook = new HSSFWorkbook(inputStream);
			} else if (ADDRESS.endsWith(".xlsx")) {
				workbook = new XSSFWorkbook(inputStream);
			}
			// IP，端口
			ServerAddress serverAddress = new ServerAddress(IP, PORT);
			List<ServerAddress> address = new ArrayList<ServerAddress>();
			address.add(serverAddress);
			// 用户名，数据库，密码
			MongoCredential credential = MongoCredential.createCredential(USERNAME, DATABASE, PASSWORD.toCharArray());
			List<MongoCredential> credentials = new ArrayList<MongoCredential>();
			credentials.add(credential);
			// 通过验证获取连接
			mongoClient = new MongoClient(address, credentials);
			// 连接到数据库
			MongoDatabase mongoDatabase = mongoClient.getDatabase(DATABASE);
			// 连接文档
			MongoCollection<Document> collection = mongoDatabase.getCollection(COLLECTION);
			System.out.println("连接成功");
			List<Document> documents = new ArrayList<Document>();
			List<String> fieldList = new ArrayList<String>();
			// 获取Excel文档中第一个表单
			Sheet sheet = workbook.getSheetAt(0);
			Row row0 = sheet.getRow(0);
			for (Cell cell : row0) {
				fieldList.add(cell.toString());
			}
			System.out.println(fieldList);
			int rows = sheet.getLastRowNum() + 1;
			int cells = fieldList.size();
			for (int i = 1; i < rows; i++) {
				System.out.println(i);
				Row row = sheet.getRow(i);
				Document document = new Document();
				for (int j = 0; j < cells; j++) {
					Cell cell = row.getCell(j);
					if (cell != null && !cell.equals("")) {
						document.append(fieldList.get(j), cell.toString());
					}
				}
				documents.add(document);
			}
			collection.insertMany(documents);
			System.out.println("导入成功");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (mongoClient != null) {
				mongoClient.close();
			}
			if (workbook != null) {
				try {
					workbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

}
