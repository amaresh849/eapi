package com.excel.custom.api.utils;

import org.apache.poi.ss.usermodel.CellType;

public enum CellDataType {

	_NONE, BOOLEAN, BLANK, ERROR, FORMULA, NUMERIC, STRING;

	public static CellType getCellType(CellDataType dataType) {
		switch (dataType) {
		case _NONE:
			return CellType._NONE;
		case BOOLEAN:
			return CellType.BOOLEAN;
		case BLANK:
			return CellType.BLANK;
		case ERROR:
			return CellType.ERROR;
		case FORMULA:
			return CellType.FORMULA;
		case NUMERIC:
			return CellType.NUMERIC;
		case STRING:
			return CellType.STRING;
		}
		return null;
	}
	
}
