package com.excel.custom.api.main.impl;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFRow;

import com.excel.custom.api.main.IRow;
import com.excel.custom.api.main.Icell;
import com.excel.custom.api.utils.CellDataType;

public final class IRowImpl implements IRow {

	private final XSSFRow xssfRow;
	private final int rIdx;

	private List<Icell> cells = new ArrayList<Icell>();

	public IRowImpl(XSSFRow xssfRow, ExcelRowStyle rowStyle) {
		this.xssfRow = xssfRow;
		this.rIdx = xssfRow.getRowNum();
		if (rowStyle != null) {
			setRowProperties(rowStyle);
		}
	}

	public IRowImpl(XSSFRow xssfRow) {
		this.xssfRow = xssfRow;
		this.rIdx = xssfRow.getRowNum();
	}

	@Override
	public int getrIdx() {
		return rIdx;
	}

	@Override
	public Icell createCell(CellDataType cellDataType, Object value,
			ExcelCellStyle cellStyle) {
		return createNewCell(cellDataType, value, cellStyle);
	}

	@Override
	public Icell createCell(CellDataType dataType, Object value) {
		return createNewCell(dataType, value, null);
	}

	private Icell createNewCell(CellDataType cellDataType, Object value,
			ExcelCellStyle cellStyle) {
		int cIdx = xssfRow.getLastCellNum();
		if (cIdx == -1)
			cIdx = 0;
		Icell cell = new IcellImpl(cIdx, rIdx, value, cellDataType,
				xssfRow.createCell(cIdx), cellStyle);
		cells.add(cell);
		return cell;
	}

	@Override
	public void setRowHeight(short height) {
		xssfRow.setHeight(height);
	}

	private void setRowProperties(ExcelRowStyle rowStyle) {
		xssfRow.setHeight((short) 1000);
		if (rowStyle.getHeightInPoints() > 0)
			xssfRow.setHeightInPoints(rowStyle.getHeightInPoints());
		xssfRow.setRowNum(rowStyle.getRowNumber());
		xssfRow.setZeroHeight(rowStyle.isZeroHeight());
		if (rowStyle.getRowStyle() != null)
			xssfRow.setRowStyle(rowStyle.getRowStyle());
	}

}
