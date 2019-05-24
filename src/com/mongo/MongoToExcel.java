package com.mongo;

import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.bson.Document;

import com.mongodb.MongoClient;
import com.mongodb.MongoCredential;
import com.mongodb.ServerAddress;
import com.mongodb.client.FindIterable;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;


public class MongoToExcel {

	private static Integer PORT = 27017;                    //端口号
	private static String IP = "localhost";                 //Ip
	private static String DATABASE = "database";            //数据库名称
	private static String USERNAME = "username";            //用户名
	private static String PASSWORD = "password";            //密码
	private static String COLLECTION = "calendar";          //文档名称
	private static String ADDRESS = "d:\\2019日历.xls";		//Excel文件所在的路径
	
	public static void main(String[] args) {
		try {
			// IP，端口
			ServerAddress serverAddress = new ServerAddress(IP, PORT);
			List<ServerAddress> address = new ArrayList<ServerAddress>();
			address.add(serverAddress);
			// 用户名，数据库，密码
			MongoCredential credential = MongoCredential.createCredential(USERNAME, DATABASE, PASSWORD.toCharArray());
			List<MongoCredential> credentials = new ArrayList<MongoCredential>();
			credentials.add(credential);
			// 通过验证获取连接
			MongoClient mongoClient = new MongoClient(address, credentials);
			// 连接到数据库
			MongoDatabase mongoDatabase = mongoClient.getDatabase(DATABASE);
			// 连接文档
			MongoCollection<Document> collection = mongoDatabase.getCollection(COLLECTION);
			// 检索所有文档
			FindIterable<Document> findIterable = collection.find();
			MongoCursor<Document> mongoCursor = findIterable.iterator();
			List<Document> documents = new ArrayList<>();
			while (mongoCursor.hasNext()) {
				documents.add(mongoCursor.next());
			}
			// 表头
			List<String> stringList = new ArrayList<>();
			for (Entry<String, Object> entry : documents.get(0).entrySet()) {
				stringList.add(entry.getKey());
			}
			// 创建HSSFWorkbook对象
			HSSFWorkbook workbook = new HSSFWorkbook();
			// 创建HSSFSheet对象
			HSSFSheet sheet = workbook.createSheet("sheet");
			// Excel表头
			HSSFRow row0 = sheet.createRow(0);
			for (int i = 0; i < stringList.size(); i++) {
				HSSFCell cell0 = row0.createCell(i);
				cell0.setCellValue(stringList.get(i));
			}
			// Excel数据
			for (int i = 0; i < documents.size(); i++) {
				HSSFRow rows = sheet.createRow(i+1);
				for (int j = 0; j < stringList.size(); j++) {
					HSSFCell cells = rows.createCell(j);
					cells.setCellValue(documents.get(i).get(stringList.get(j)).toString());
				}
			}
			// 输出文件
			FileOutputStream outputStream = new FileOutputStream(ADDRESS);
			workbook.write(outputStream);
			outputStream.flush();
			System.out.println("导出成功");
		} catch (Exception e) {
			System.err.println(e.getClass().getName() + ": " + e.getMessage());
		}
	}
	
}
