package com.excel.custom.api.main;

import java.io.IOException;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.excel.custom.api.main.impl.ExcelCellStyle;

public interface IWorkbook {

	public abstract IWorkbook getWorkbook();

	public abstract String getName();

	public abstract ISheet createSheet(String sheetName);

	public abstract List<ISheet> getSheets();

	public abstract void write() throws IOException;

	public abstract ExcelCellStyle createCellStyle();

	public abstract XSSFWorkbook getXssfWorkbook();
}
