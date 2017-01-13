package com.excel.custom.api.utils;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.RichTextString;

public enum CellValue {
	
	STRING,DOUBLE,BOOLEAN,CALENDAR,RICHTEXTSTRING,DATE,NUMERIC;
	
	public static String getStringValue(Object cellValue) {
		return cellValue.toString();
	} 
	
	public static boolean getBooleanValue(Object value) {
		return (Boolean)value;
	}
	
	public static Date getDateValue(Object value) {
		return (Date)value;
	}
	
	public static Calendar getCalendarValue(Object value) {
		return(Calendar)value;
	}
	
	public static double getDoubleValue(Object value) {
		return (Double)value;
	}
	
	public static RichTextString getRichTextString(Object value) {
		return (RichTextString)value;
	}
	
	public static Integer getInteger(Object value) {
		return (Integer)value;
	}
	
}
