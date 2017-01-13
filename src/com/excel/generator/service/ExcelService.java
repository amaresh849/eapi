package com.excel.generator.service;

import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import com.excel.custom.api.main.IRow;
import com.excel.custom.api.main.ISheet;
import com.excel.custom.api.main.IWorkbook;
import com.excel.custom.api.main.impl.ExcelCellStyle;
import com.excel.custom.api.main.impl.IWorkbookImpl;
import com.excel.custom.api.utils.CellDataType;
import com.excel.generator.helper.DataProvider;
import com.excel.generator.helper.EnHeaderColumnNames;
import com.excel.generator.helper.ProductDto;

public class ExcelService {

	public static void main(String[] args) throws IOException {
		ExcelService service = new ExcelService();
		service.generateExcel();
	}

	public void generateExcel() throws IOException {

		IWorkbook workbook = new IWorkbookImpl("Workbook.xlsx");
		ISheet sheet = workbook.createSheet("Sheet");
		EnHeaderColumnNames[] header = EnHeaderColumnNames.values();
		ExcelCellStyle headerCellStyle = setHeaderCellProperties(workbook);
		setHeaderRowProperties(sheet);
		int cIdx = -1;
		for (EnHeaderColumnNames enHeaderColumnName : header) {
			sheet.setColumnWidth(++cIdx, 10 * 500);
			sheet.createHeaderRow(enHeaderColumnName.toString(),
					headerCellStyle);
		}
		DataProvider dataProvider = new DataProvider();
		List<ProductDto> productData = dataProvider.getProductData();
		ExcelCellStyle dataCellStyle = setDataCellProperties(workbook);
		for (ProductDto productDto : productData) {
			IRow row = sheet.createDataRow();
			row.createCell(CellDataType.NUMERIC, productDto.getProductId(),
					dataCellStyle);
			row.createCell(CellDataType.STRING, productDto.getProductName(),
					dataCellStyle);
			row.createCell(CellDataType.NUMERIC, productDto.getProductPrice(),
					dataCellStyle);
			row.createCell(CellDataType.NUMERIC, productDto.getQuantity(),
					dataCellStyle);
			// row.createCell(CellDataType.NUMERIC,
			// productDto.getManufacturingDate(), dataCellStyle);
		}
		workbook.write();
	}

	public ExcelCellStyle setDataCellProperties(IWorkbook workbook) {
		ExcelCellStyle dataCellStyle = workbook.createCellStyle();
		dataCellStyle.setFillForegroundColor(IndexedColors.GOLD.index);
		dataCellStyle.setBorderBottom(BorderStyle.THIN);
		dataCellStyle.setBorderLeft(BorderStyle.THIN);
		dataCellStyle.setBorderRightEnum(BorderStyle.THIN);
		dataCellStyle.setBorderTopEnum(BorderStyle.THIN);
		dataCellStyle.setFillPatternEnum(FillPatternType.SOLID_FOREGROUND);
		return dataCellStyle;
	}

	public void setHeaderRowProperties(ISheet sheet) {
		sheet.setHeaderRowHeight((short) 1000);
	}

	public ExcelCellStyle setHeaderCellProperties(IWorkbook workbook) {
		ExcelCellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFillForegroundColor(IndexedColors.GREEN.index);
		headerCellStyle.setBorderBottom(BorderStyle.THIN);
		headerCellStyle.setBorderLeft(BorderStyle.THIN);
		headerCellStyle.setBorderRightEnum(BorderStyle.THIN);
		headerCellStyle.setBorderTopEnum(BorderStyle.THIN);
		headerCellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
		return headerCellStyle;
	}

}
