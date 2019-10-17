package com.teclan.poi.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Objects {
	private static final Logger LOGGER = LoggerFactory.getLogger(Objects.class);

	public static boolean isNull(Object value) {
		return value == null;
	}

	public static boolean isNotNull(Object value) {
		return !isNull(value);
	}

	public static boolean isNotNullString(Object value) {
		return !isNullString(value);
	}

	/**
	 * 字符串是否是 null 或者 ""
	 * 
	 * @param value
	 * @return
	 */
	public static boolean isNullString(Object value) {
		return (value == null || "".equals(value.toString().trim()));
	}
}
