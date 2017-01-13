package com.excel.custom.api.test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFCellStyle cellstyle = workbook.createCellStyle();
		cellstyle.setFillForegroundColor(IndexedColors.BLUE.index);
		cellstyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		XSSFSheet sheet = workbook.createSheet("Sheet");
		XSSFRow row = sheet.createRow(0);
		XSSFCell cell = row.createCell(0);
		cell.setCellType(CellType.STRING);
		cell.setCellValue("abc");
		cell.setCellStyle(cellstyle);
		System.out.println(cell.getCellStyle().getFillForegroundColor());
		System.out.println(cell.getCellStyle().getFillPattern());
		FileOutputStream os = new FileOutputStream("Workbook.xlsx");
		workbook.write(os);
		workbook.close();
	}

}
