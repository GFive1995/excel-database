package com.mongo;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.compress.utils.Lists;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.bson.Document;

import com.mongodb.MongoClient;
import com.mongodb.MongoCredential;
import com.mongodb.ServerAddress;
import com.mongodb.client.FindIterable;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;
import com.util.ExcelUtil;


public class MongoToExcel {

	private static Integer PORT = 27017; 				// 端口号
	private static String IP = "localhost"; 			// Ip
	private static String DATABASE = "database"; 		// 数据库名称
	private static String USERNAME = "username"; 		// 用户名
	private static String PASSWORD = "password"; 		// 密码
	private static String COLLECTION = "calendar"; 		// 文档名称
	private static String PATH = "d:\\2019日历.xlsx"; 	// Excel文件所在的路径
	private static String FILE_NAME = "2019日历.xlsx";	// 文件名
	
	public static void main(String[] args) {
		MongoClient mongoClient = null;
		HSSFWorkbook workbook = null;
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
			mongoClient = new MongoClient(address, credentials);
			// 连接到数据库
			MongoDatabase mongoDatabase = mongoClient.getDatabase(DATABASE);
			// 连接文档
			MongoCollection<Document> collection = mongoDatabase.getCollection(COLLECTION);
			// 检索所有文档
			FindIterable<Document> findIterable = collection.find();
			MongoCursor<Document> mongoCursor = findIterable.iterator();
			List<Map<String, Object>> dataList = Lists.newArrayList();
			List<String> fieldList = new ArrayList<>();
			while (mongoCursor.hasNext()) {
				Document document = mongoCursor.next();
				if (fieldList==null || fieldList.size()==0) {
					for (Entry<String, Object> entry : document.entrySet()) {
						fieldList.add(entry.getKey());
					}
				} 
				Map<String, Object> paraMap = new HashMap<String, Object>();
				for (int i = 0; i < fieldList.size(); i++) {
					paraMap.put(fieldList.get(i), document.get(fieldList.get(i)));
				}
				dataList.add(paraMap);
			}
			ExcelUtil.downloadExcel(FILE_NAME, PATH, fieldList, fieldList, dataList);
			System.out.println("导出成功");
		} catch (Exception e) {
			System.err.println(e.getClass().getName() + ": " + e.getMessage());
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
