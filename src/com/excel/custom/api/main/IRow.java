package com.excel.custom.api.main;

import com.excel.custom.api.main.impl.ExcelCellStyle;
import com.excel.custom.api.utils.CellDataType;

public interface IRow {

	public abstract int getrIdx();

	public abstract Icell createCell(CellDataType cellDataType, Object value,
			ExcelCellStyle cellStyle);

	public abstract Icell createCell(CellDataType dataType, Object value);
	
	void setRowHeight(short height);
}
