package com.excel.custom.api.main.impl;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import com.excel.custom.api.main.Icell;
import com.excel.custom.api.utils.CellDataType;
import com.excel.custom.api.utils.CellValue;

public final class IcellImpl implements Icell {

	private final int cIdx;
	private final int rIdx;
	private final Object value;
	private final CellDataType dataType;
	private final ExcelCellStyle cellStyle;

	private XSSFCell xssfCell;
	private XSSFCellStyle xssfCellStyle;

	public IcellImpl(int cIdx, int rIdx, Object value, CellDataType dataType,
			XSSFCell xssfCell, ExcelCellStyle cellStyle) {
		this.cIdx = cIdx;
		this.rIdx = rIdx;
		this.value = value;
		this.dataType = dataType;
		this.xssfCell = xssfCell;
		this.cellStyle = cellStyle;
		this.xssfCellStyle = xssfCell.getSheet().getWorkbook()
				.createCellStyle();
		setCellProperties();
	}

	private void setCellValue(Object value, XSSFCell xssfCell) {
		if (value instanceof String) {
			xssfCell.setCellValue(CellValue.getStringValue(value));
		} else if (value instanceof Boolean) {
			xssfCell.setCellValue(CellValue.getBooleanValue(value));
		} else if (value instanceof Date) {
			xssfCell.setCellValue(CellValue.getDateValue(value));
		} else if (value instanceof Calendar) {
			xssfCell.setCellValue(CellValue.getCalendarValue(value));
		} else if (value instanceof RichTextString) {
			xssfCell.setCellValue(CellValue.getRichTextString(value));
		} else if (value instanceof Double) {
			xssfCell.setCellValue(CellValue.getDoubleValue(value));
		} else if (value instanceof Integer) {
			xssfCell.setCellValue(CellValue.getInteger(value));
		}
	}

	private void setCellStyle() {
		if (cellStyle != null) {
			if (cellStyle.getAlignment() > 0)
				xssfCellStyle.setAlignment(cellStyle.getAlignment());
			if (cellStyle.getAlignmentEnum() != null) {
				xssfCellStyle.setAlignment(cellStyle.getAlignmentEnum());
			}
			if (cellStyle.getBorderBottomEnum() != null) {
				xssfCellStyle.setBorderBottom(cellStyle.getBorderBottomEnum());
			} else {
				xssfCellStyle.setBorderBottom(BorderStyle.NONE);
			}
			if (cellStyle.getBorderBottom() > 0)
				xssfCellStyle.setBorderBottom(cellStyle.getBorderBottom());
			// xssfCellStyle.setBorderColor(cellStyle.get, color)
			if (cellStyle.getBorderLeft() > 0)
				xssfCellStyle.setBorderLeft(cellStyle.getBorderLeft());
			if (cellStyle.getBorderLeftEnum() != null) {
				xssfCellStyle.setBorderLeft(cellStyle.getBorderLeftEnum());
			} else {
				xssfCellStyle.setBorderLeft(BorderStyle.NONE);
			}
			if (cellStyle.getBorderRight() > 0)
				xssfCellStyle.setBorderRight(cellStyle.getBorderRight());
			if (cellStyle.getBorderRightEnum() != null) {
				xssfCellStyle.setBorderRight(cellStyle.getBorderRightEnum());
			} else {
				xssfCellStyle.setBorderRight(BorderStyle.NONE);
			}
			if (cellStyle.getBorderTop() > 0)
				xssfCellStyle.setBorderTop(cellStyle.getBorderTop());
			if (cellStyle.getBorderTopEnum() != null) {
				xssfCellStyle.setBorderTop(cellStyle.getBorderTopEnum());
			} else {
				xssfCellStyle.setBorderTop(BorderStyle.NONE);
			}
			if (cellStyle.getBottomBorderColor() > 0)
				xssfCellStyle.setBottomBorderColor(cellStyle
						.getBottomBorderColor());
			// if (cellStyle.getBorderBottomEnum() != null)
			// xssfCellStyle.setBottomBorderColor(cellStyle
			// .getBottomBorderXSSFColor());
			if (cellStyle.getDataFormat() > 0)
				xssfCellStyle.setDataFormat(cellStyle.getDataFormat());
			// xssfCellStyle.setDataFormat(cellStyle.getd)
			if (cellStyle.getFillBackgroundColor() > 0)
				xssfCellStyle.setFillBackgroundColor(cellStyle
						.getFillBackgroundColor());
			if (cellStyle.getFillBackgroundXSSFColor() != null) {
				xssfCellStyle.setFillBackgroundColor(cellStyle
						.getFillBackgroundXSSFColor());
			}
			if (cellStyle.getFillForegroundColor() > 0)
				xssfCellStyle.setFillForegroundColor(cellStyle
						.getFillForegroundColor());
			if (cellStyle.getFillForegroundXSSFColor() != null) {
				xssfCellStyle.setFillForegroundColor(cellStyle
						.getFillForegroundXSSFColor());
			}
			if (cellStyle.getFillPattern() > 0) {
				xssfCellStyle.setFillPattern(cellStyle.getFillPattern());
			}
			if (cellStyle.getFillPatternEnum() != null) {
				xssfCellStyle.setFillPattern(cellStyle.getFillPatternEnum());
			}
			if (cellStyle.getFont() != null) {
				xssfCellStyle.setFont(cellStyle.getFont());
			}
			xssfCellStyle.setHidden(cellStyle.isHidden());
			if (cellStyle.getIndention() > 0)
				xssfCellStyle.setIndention(cellStyle.getIndention());
			if (cellStyle.getLeftBorderColor() > 0)
				xssfCellStyle
						.setLeftBorderColor(cellStyle.getLeftBorderColor());
			if (cellStyle.getLeftBorderXSSFColor() != null) {
				xssfCellStyle.setLeftBorderColor(cellStyle
						.getLeftBorderXSSFColor());
			}
			xssfCellStyle.setLocked(cellStyle.isLocked());
			if (cellStyle.getRightBorderColor() > 0)
				xssfCellStyle.setRightBorderColor(cellStyle
						.getRightBorderColor());
			if (cellStyle.getRightBorderXSSFColor() != null) {
				xssfCellStyle.setRightBorderColor(cellStyle
						.getRightBorderXSSFColor());
			}
			// if (cellStyle.getRotation() > 0)
			// xssfCellStyle.setRotation(cellStyle.getRotation());
			// // xssfCellStyle.setShrinkToFit(cellStyle.get)
			xssfCellStyle.setTopBorderColor((short) 8);
			if (cellStyle.getTopBorderXSSFColor() != null) {
				xssfCellStyle.setTopBorderColor(cellStyle
						.getTopBorderXSSFColor());
			}
			if (cellStyle.getVerticalAlignment() > 0)
				xssfCellStyle.setVerticalAlignment(cellStyle
						.getVerticalAlignment());
			if (cellStyle.getVerticalAlignmentEnum() != null) {
				xssfCellStyle.setVerticalAlignment(cellStyle
						.getVerticalAlignmentEnum());
			}
			xssfCellStyle.setWrapText(cellStyle.isWrapText());
			xssfCell.setCellStyle(xssfCellStyle);
		}
	}

	private void setCellProperties() {
		setCellValue(value, xssfCell);
		xssfCell.setCellType(CellDataType.getCellType(dataType));
		setCellStyle();
	}

	@Override
	public int getcIdx() {
		return cIdx;
	}

	@Override
	public int getrIdx() {
		return rIdx;
	}

	@Override
	public Object getValue() {
		return value;
	}

	@Override
	public CellDataType getDataType() {
		return dataType;
	}

	@Override
	public XSSFCell getXssfCell() {
		return xssfCell;
	}

	@Override
	public ExcelCellStyle getCellStyle() {
		return cellStyle;
	}

	@Override
	public XSSFCellStyle getXssfCellStyle() {
		return xssfCellStyle;
	}

}
