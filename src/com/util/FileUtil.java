package com.util;

import java.io.File;
import java.io.IOException;


public class FileUtil {

	/**
	 * 
	 * 方法描述:判断路径文件是否存在，如果不存在则创建
	 *
	 * @param path
	 * 
	 * @author wangcy
	 * @date 2019年6月26日 下午4:44:51
	 */
	public static void mkdir(String path) {
        try {
        	File file = new File(path);
        	if (!file.getParentFile().exists()) { 
        		file.getParentFile().mkdirs();
        	}
        	if (file.exists()) { 
        		file.delete();
        	}
			file.createNewFile();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
}
