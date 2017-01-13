package com.excel.custom.api.main.impl;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import com.excel.custom.api.main.IRow;
import com.excel.custom.api.main.ISheet;
import com.excel.custom.api.utils.CellDataType;

public final class ISheetImpl implements ISheet {

	private final String sheetName;
	private XSSFSheet xssfSheet;
	private List<IRow> rows = new ArrayList<IRow>();

	private List<ExcelRowStyle> rowStyles = new ArrayList<ExcelRowStyle>();

	public ISheetImpl(XSSFSheet xssfSheet) {
		this.xssfSheet = xssfSheet;
		this.sheetName = xssfSheet.getSheetName();
	}

	@Override
	public String getSheetName() {
		return sheetName;
	}
	
	@Override
	public XSSFSheet getXssfSheet() {
		return xssfSheet;
	}

	public List<IRow> getRows() {
		return rows;
	}

	@Override
	public IRow createHeaderRow(String columnName, ExcelCellStyle cellStyle) {
		IRow row;
		int headerRowNum = 0;
		if (rows == null || rows.isEmpty()) {
			if (!rowStyles.isEmpty())
				row = new IRowImpl(xssfSheet.createRow(headerRowNum),
						rowStyles.get(headerRowNum));
			else
				row = new IRowImpl(xssfSheet.createRow(headerRowNum));
			row.createCell(CellDataType.STRING, columnName, cellStyle);
			rows.add(row);
			return row;
		} else {
			row = rows.get(headerRowNum);
			row.createCell(CellDataType.STRING, columnName, cellStyle);
			rows.add(row);
			return row;
		}
	}

	@Override
	public IRow createDataRow() {
		XSSFRow xssfRow = getRow();
		IRow row;
		if (!rowStyles.isEmpty()
				&& (rowStyles.size() >= xssfRow.getRowNum() + 1))
			row = new IRowImpl(xssfRow, rowStyles.get(xssfRow.getRowNum()));
		else
			row = new IRowImpl(xssfRow);
		rows.add(row);
		return row;
	}

	private XSSFRow getRow() {
		int rIdx = 0;
		XSSFRow xssfRow = null;
		xssfRow = xssfSheet.getRow(rIdx);
		while (xssfRow != null) {
			rIdx++;
			xssfRow = xssfSheet.getRow(rIdx);
		}
		return xssfSheet.createRow(rIdx);
	}

	@Override
	public void setColumnWidth(int cIdx, int width) {
		xssfSheet.setColumnWidth(cIdx, width);
	}

	@Override
	public void setHeaderRowHeight(short height) {
		int headerRowIndex = 0;
		ExcelRowStyle rowStyle = new ExcelRowStyle(headerRowIndex);
		rowStyle.setHeight(height);
		rowStyles.add(rowStyle);
	}

	@Override
	public void setDataRowHeight(int rowIndex, short height) {
		ExcelRowStyle rowStyle = new ExcelRowStyle(rowIndex);
		rowStyle.setHeight(height);
		rowStyles.add(rowStyle);
	}

	@Override
	public void createFreezePane(int colSplit, int rowSplit) {
		xssfSheet.createFreezePane(colSplit, rowSplit);
	}

	@Override
	public void createFreezePane(int colSplit, int rowSplit,
			int leftmostColumn, int topRow) {
		xssfSheet.createFreezePane(colSplit, rowSplit, leftmostColumn, topRow);
	}

}
