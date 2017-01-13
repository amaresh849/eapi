package com.excel.generator.service;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.excel.generator.helper.EnHeaderColumnNames;

public class Test {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("defaultproperties1");
		XSSFRow headerRow = sheet.createRow(0);
		EnHeaderColumnNames[] headerNames = EnHeaderColumnNames.values();
		int cIdx = 0;
		XSSFCell cell = null;
		XSSFCellStyle cs = workbook.createCellStyle();
		for (EnHeaderColumnNames enHeaderColumnName : headerNames) {
			cell = headerRow.createCell(cIdx++, CellType.STRING);
			cell.setCellValue(enHeaderColumnName.toString());
			cs.setFillForegroundColor(IndexedColors.WHITE.index);
			cs.setFillPattern(XSSFCellStyle.THIN_BACKWARD_DIAG);
			cell.setCellStyle(cs);
		}
		FileOutputStream os = new FileOutputStream(
				"C:/Users/win/Desktop/Default.xlsx");
		workbook.write(os);
		workbook.close();
	}
}
