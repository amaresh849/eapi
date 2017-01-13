package com.excel.custom.api.main;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import com.excel.custom.api.main.impl.ExcelCellStyle;
import com.excel.custom.api.utils.CellDataType;

public interface Icell {

	public abstract int getcIdx();

	public abstract int getrIdx();

	public abstract Object getValue();

	public abstract CellDataType getDataType();

	public abstract XSSFCell getXssfCell();

	public abstract ExcelCellStyle getCellStyle();

	public abstract XSSFCellStyle getXssfCellStyle();

}
