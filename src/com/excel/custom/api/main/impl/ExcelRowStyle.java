package com.excel.custom.api.main.impl;

import org.apache.poi.ss.usermodel.CellStyle;

public class ExcelRowStyle {

	private int rowIndex;

	private short height;
	private float heightInPoints;
	private int rowNumber;
	private CellStyle rowStyle;
	private boolean zeroHeight;

	public ExcelRowStyle(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public short getHeight() {
		return height;
	}

	public void setHeight(short height) {
		this.height = height;
	}

	public float getHeightInPoints() {
		return heightInPoints;
	}

	public void setHeightInPoints(float heightInPoints) {
		this.heightInPoints = heightInPoints;
	}

	public int getRowNumber() {
		return rowNumber;
	}

	public void setRowNumber(int rowNumber) {
		this.rowNumber = rowNumber;
	}

	public CellStyle getRowStyle() {
		return rowStyle;
	}

	public void setRowStyle(CellStyle rowStyle) {
		this.rowStyle = rowStyle;
	}

	public boolean isZeroHeight() {
		return zeroHeight;
	}

	public void setZeroHeight(boolean zeroHeight) {
		this.zeroHeight = zeroHeight;
	}

}
