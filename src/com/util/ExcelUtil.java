package com.util;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;


public class ExcelUtil {

	/**
	 * 
	 * 方法描述:将dataList数据保存到Excel中
	 *
	 * @param excelName	文件名
	 * @param path		文件路径
	 * @param headList	表头
	 * @param fieldList	表头对应字段
	 * @param dataList	数据
	 * 
	 * @author wangcy
	 * @date 2019年6月26日 下午4:35:38
	 */
	public static void downloadExcel(String excelName, String path, List<String> headList, List<String> fieldList, List<Map<String, Object>> dataList) {
		SXSSFWorkbook workbook = null;
		FileOutputStream outputStream = null;
		try {
			FileUtil.mkdir(path);
			workbook = new SXSSFWorkbook();
			Sheet sheet = workbook.createSheet(excelName);
			Row row_0 = sheet.createRow(0);
			for (int i = 0; i < headList.size(); i++) {
				Cell cell_i = row_0.createCell(i);
				cell_i.setCellValue(headList.get(i));
			}
			if (dataList!=null && dataList.size()!=0) {
				for (int i = 0; i < dataList.size(); i++) {
    				Row row = sheet.createRow(i+1);
    				for (int j = 0; j < fieldList.size(); j++) {
    					Cell cell = row.createCell(j);
    					cell.setCellValue(ObjectUtils.castString(dataList.get(i).get(fieldList.get(j)), "") );
    				}
    			}
			}
			outputStream = new FileOutputStream(path);
			workbook.write(outputStream);
			outputStream.flush();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (workbook!=null) {
				try {
					workbook.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (outputStream!=null) {
				try {
					outputStream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			
		}
	}
	
}
