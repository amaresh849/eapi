package com.excel.custom.api.main.impl;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.excel.custom.api.main.ISheet;
import com.excel.custom.api.main.IWorkbook;

public final class IWorkbookImpl implements IWorkbook {

	private XSSFWorkbook xssfWorkbook;
	private final String name;
	private List<ISheet> sheets = new ArrayList<ISheet>();

	public IWorkbookImpl(String name) {
		this.xssfWorkbook = new XSSFWorkbook();
		this.name = name;
	}

	@Override
	public XSSFWorkbook getXssfWorkbook() {
		return xssfWorkbook;
	}

	@Override
	public IWorkbook getWorkbook() {
		return this;
	}

	@Override
	public String getName() {
		return name;
	}

	@Override
	public ISheet createSheet(String sheetName) {
		ISheet sheet = new ISheetImpl(xssfWorkbook.createSheet(sheetName));
		sheets.add(sheet);
		return sheet;
	}

	@Override
	public List<ISheet> getSheets() {
		return sheets;
	}

	@Override
	public ExcelCellStyle createCellStyle() {
		return new ExcelCellStyle(xssfWorkbook.createCellStyle());
	}

	@Override
	public void write() throws IOException {
		FileOutputStream fos = new FileOutputStream(getName());
		xssfWorkbook.write(fos);
		xssfWorkbook.close();
	}

}
