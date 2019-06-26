package com.util;

import java.math.BigDecimal;
import java.util.Date;

public class ObjectUtils {
	public static String castString(Object o) {
		return ObjectUtils.castString(o, null);
	}

	public static String castString(Object o, String defaultValue) {
		if (o == null) {
			return defaultValue;
		}
		return o.toString();
	}

	public static Long castLong(Object obj) {
		return ObjectUtils.castLong(obj, null);
	}

	public static Long castLong(Object obj, Long defaultValue) {
		if (obj == null || "".equals(obj)) {
			return defaultValue;
		}
		return Long.valueOf(obj.toString());
	}

	public static Integer castInteger(Object obj) {
		return ObjectUtils.castInteger(obj, null);
	}

	public static Integer castInteger(Object obj, Integer defaultValue) {
		if (obj == null || "".equals(obj)) {
			return defaultValue;
		}
		return Integer.valueOf(obj.toString());
	}

	public static Float castFloat(Object obj) {
		return ObjectUtils.castFloat(obj, null);
	}

	public static Float castFloat(Object obj, Float defaultValue) {
		if (obj == null || "".equals(obj)) {
			return defaultValue;
		}
		return Float.valueOf(obj.toString());
	}

	public static Double castDouble(Object obj) {
		return ObjectUtils.castDouble(obj, null);
	}

	public static Double castDouble(Object obj, Double defaultValue) {
		if (obj == null || "".equals(obj)) {
			return defaultValue;
		}
		return Double.valueOf(obj.toString());
	}

	public static Date castDate(Object obj) {
		return ObjectUtils.castDate(obj, null);
	}

	public static Date castDate(Object obj, Date defaultValue) {
		if (obj == null || "".equals(obj)) {
			return defaultValue;
		}
		if (obj instanceof Date) {
			return (Date) obj;
		}
		return defaultValue;
	}

	public static Double getDouble(Double a, Integer num) {
		BigDecimal bg = new BigDecimal(a);
		return bg.setScale((int) num, 4).doubleValue();
	}
}
