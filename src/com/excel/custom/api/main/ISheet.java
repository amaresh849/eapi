package com.excel.custom.api.main;

import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.excel.custom.api.main.impl.ExcelCellStyle;

public interface ISheet {

	public abstract String getSheetName();

	public abstract IRow createDataRow();

	public abstract IRow createHeaderRow(String columnName,
			ExcelCellStyle cellStyle);

	public abstract void setColumnWidth(int cIdx, int width);

	public abstract void setHeaderRowHeight(short height);

	public abstract void setDataRowHeight(int rowIndex, short height);

	public abstract void createFreezePane(int colSplit, int rowSplit);

	public abstract void createFreezePane(int colSplit, int rowSplit,
			int leftmostColumn, int topRow);

	public abstract XSSFSheet getXssfSheet();

}
