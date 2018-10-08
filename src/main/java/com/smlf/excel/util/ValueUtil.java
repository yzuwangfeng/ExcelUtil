package com.smlf.excel.util;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;

import com.smlf.excel.annotation.DataField;

public class ValueUtil {
	
	/**
	 * 参数格式化为String
	 *
	 * @param field
	 * @param value
	 * @return
	 */
	public static String formatValue(Field field, Object value) {
		Class<?> fieldType = field.getType();

		DataField excelField = field.getAnnotation(DataField.class);
		if(value==null) {
			return null;
		}

		if (Boolean.class.equals(fieldType) || Boolean.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (String.class.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Short.class.equals(fieldType) || Short.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Integer.class.equals(fieldType) || Integer.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Long.class.equals(fieldType) || Long.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Float.class.equals(fieldType) || Float.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Double.class.equals(fieldType) || Double.TYPE.equals(fieldType)) {
			return String.valueOf(value);
		} else if (Date.class.equals(fieldType)) {
			String datePattern = "yyyy-MM-dd HH:mm:ss";
			if (excelField != null && excelField.dateformat()!=null) {
				datePattern = excelField.dateformat();
			}
			SimpleDateFormat dateFormat = new SimpleDateFormat(datePattern);
			return dateFormat.format(value);
		} else {
			throw new RuntimeException("request illeagal type, type must be Integer not int Long not long etc, type=" + fieldType);
		}
	}

}
