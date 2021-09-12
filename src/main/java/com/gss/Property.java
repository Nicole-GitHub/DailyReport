package com.gss;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

public class Property {
	protected static Map<String, String> getProperties(String path) {
		Map<String, String> map = new HashMap<String, String>();
		Properties prop = new Properties();

		try (FileInputStream fis = new FileInputStream(path + "config.properties")) {

			// 加載屬性
			prop.load(fis);

			// 取得所有鍵的列舉
			Enumeration<?> e = prop.propertyNames();
			while (e.hasMoreElements()) {
				// 取得下一個鍵
				String key = (String) e.nextElement();
				// 取得 properties 屬性值
				String value = prop.getProperty(key, "搜尋不到 " + key);
//				System.out.println(key + " = " + value);
				map.put(key, value);
			}

		} catch (IOException e) {
			System.out.println("getProperty Error:" + e.getMessage());
		}

		return map;
	}
}
