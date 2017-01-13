package com.excel.custom.api.main.impl;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

public class ExcelCellStyle {

	private final XSSFCellStyle xssfCellStyle;

	ExcelCellStyle(XSSFCellStyle xssfCellStyle) {
		this.xssfCellStyle = xssfCellStyle;
	}

	private short alignment;
	private HorizontalAlignment alignmentEnum;
	private short borderBottom;
	private BorderStyle borderBottomEnum;
	private short borderLeft;
	private BorderStyle borderLeftEnum;
	private short borderRight;
	private BorderStyle borderRightEnum;
	private short borderTop;
	private BorderStyle borderTopEnum;
	private short bottomBorderColor;
	private XSSFColor bottomBorderXSSFColor;
	private short dataFormat;
	private String dataFormatString;
	private short fillBackgroundColor;
	private XSSFColor fillBackgroundColorColor;
	private XSSFColor fillBackgroundXSSFColor;
	private short fillForegroundColor;
	private short fillForegroundColorColor;
	private XSSFColor fillForegroundXSSFColor;
	private short fillPattern;
	private FillPatternType fillPatternEnum;
	private XSSFFont font;
	private short fontIndex;
	private boolean hidden;
	private short indention;
	private short index;
	private short leftBorderColor;
	private XSSFColor leftBorderXSSFColor;
	private boolean locked;
	private short rightBorderColor;
	private XSSFColor rightBorderXSSFColor;
	private short rotation;
	private short topBorderColor;
	private XSSFColor topBorderXSSFColor;
	private short verticalAlignment;
	private VerticalAlignment verticalAlignmentEnum;
	private boolean wrapText;

	public XSSFCellStyle getXssfCellStyle() {
		return xssfCellStyle;
	}

	public short getAlignment() {
		return alignment;
	}

	@Deprecated
	public void setAlignment(short alignment) {
		if (xssfCellStyle != null && alignment > 0) {
			this.alignment = alignment;
			xssfCellStyle.setAlignment(alignment);
		}
	}

	public HorizontalAlignment getAlignmentEnum() {
		return alignmentEnum;
	}

	public void setAlignmentEnum(HorizontalAlignment alignmentEnum) {
		if (xssfCellStyle != null) {
			this.alignmentEnum = alignmentEnum;
			xssfCellStyle.setAlignment(alignmentEnum);
		}
	}

	public short getBorderBottom() {
		return borderBottom;
	}

	@Deprecated
	public void setBorderBottom(short borderBottom) {
		if (xssfCellStyle != null && borderBottom > 0) {
			this.borderBottom = borderBottom;
			xssfCellStyle.setBorderBottom(borderBottom);
		}
	}

	public BorderStyle getBorderBottomEnum() {
		return borderBottomEnum;
	}

	public void setBorderBottom(BorderStyle borderBottomEnum) {
		if (xssfCellStyle != null) {
			this.borderBottomEnum = borderBottomEnum;
			xssfCellStyle.setBorderBottom(borderBottomEnum);
		}
	}

	public short getBorderLeft() {
		return borderLeft;
	}

	@Deprecated
	public void setBorderLeft(short borderLeft) {
		if (xssfCellStyle != null && borderLeft > 0) {
			this.borderLeft = borderLeft;
			xssfCellStyle.setBorderLeft(borderLeft);
		}
	}

	public BorderStyle getBorderLeftEnum() {
		return borderLeftEnum;
	}

	public void setBorderLeft(BorderStyle borderLeftEnum) {
		if (xssfCellStyle != null) {
			this.borderLeftEnum = borderLeftEnum;
			xssfCellStyle.setBorderLeft(borderLeftEnum);
		}
	}

	public short getBorderRight() {
		return borderRight;
	}

	@Deprecated
	public void setBorderRight(short borderRight) {
		if (xssfCellStyle != null && borderRight > 0) {
			this.borderRight = borderRight;
			xssfCellStyle.setBorderRight(borderRight);
		}
	}

	public BorderStyle getBorderRightEnum() {
		return borderRightEnum;
	}

	public void setBorderRightEnum(BorderStyle borderRightEnum) {
		if (xssfCellStyle != null) {
			this.borderRightEnum = borderRightEnum;
			xssfCellStyle.setBorderRight(borderRightEnum);
		}
	}

	public short getBorderTop() {
		return borderTop;
	}

	@Deprecated
	public void setBorderTop(short borderTop) {
		if (xssfCellStyle != null && borderTop > 0) {
			this.borderTop = borderTop;
			xssfCellStyle.setBorderTop(borderTop);
		}
	}

	public BorderStyle getBorderTopEnum() {
		return borderTopEnum;
	}

	public void setBorderTopEnum(BorderStyle borderTopEnum) {
		if (xssfCellStyle != null) {
			this.borderTopEnum = borderTopEnum;
			xssfCellStyle.setBorderTop(borderTopEnum);
		}
	}

	public short getBottomBorderColor() {
		return bottomBorderColor;
	}

	public void setBottomBorderColor(short bottomBorderColor) {
		if (xssfCellStyle != null && bottomBorderColor > 0) {
			this.bottomBorderColor = bottomBorderColor;
			xssfCellStyle.setBottomBorderColor(bottomBorderColor);
		}
	}

	public XSSFColor getBottomBorderXSSFColor() {
		return bottomBorderXSSFColor;
	}

	public void setBottomBorderXSSFColor(XSSFColor bottomBorderXSSFColor) {
		if (xssfCellStyle != null) {
			this.bottomBorderXSSFColor = bottomBorderXSSFColor;
			xssfCellStyle.setBottomBorderColor(bottomBorderXSSFColor);
		}
	}

	public short getDataFormat() {
		return dataFormat;
	}

	public void setDataFormat(short dataFormat) {
		if (xssfCellStyle != null && dataFormat > 0) {
			this.dataFormat = dataFormat;
			xssfCellStyle.setDataFormat(dataFormat);
		}
	}

	public String getDataFormatString() {
		return dataFormatString;
	}

	public void setDataFormatString(String dataFormatString) {
		if (xssfCellStyle != null && !dataFormatString.isEmpty()) {
			this.dataFormatString = dataFormatString;
			// xssfCellStyle.setdatafor
		}
	}

	public short getFillBackgroundColor() {
		return fillBackgroundColor;
	}

	public void setFillBackgroundColor(short fillBackgroundColor) {
		if (xssfCellStyle != null && fillBackgroundColor > 0) {
			this.fillBackgroundColor = fillBackgroundColor;
			xssfCellStyle.setFillBackgroundColor(fillBackgroundColor);
		}
	}

	public XSSFColor getFillBackgroundColorColor() {
		return fillBackgroundColorColor;
	}

	public void setFillBackgroundColorColor(XSSFColor fillBackgroundColorColor) {
		if (xssfCellStyle != null) {
			this.fillBackgroundColorColor = fillBackgroundColorColor;
			xssfCellStyle.setFillBackgroundColor(fillBackgroundColorColor);
		}
	}

	public XSSFColor getFillBackgroundXSSFColor() {
		return fillBackgroundXSSFColor;
	}

	public void setFillBackgroundXSSFColor(XSSFColor fillBackgroundXSSFColor) {
		if (xssfCellStyle != null) {
			this.fillBackgroundXSSFColor = fillBackgroundXSSFColor;
			xssfCellStyle.setFillBackgroundColor(fillBackgroundXSSFColor);
		}
	}

	public short getFillForegroundColor() {
		return fillForegroundColor;
	}

	public void setFillForegroundColor(short fillForegroundColor) {
		if (xssfCellStyle != null && fillForegroundColor > 0) {
			this.fillForegroundColor = fillForegroundColor;
			xssfCellStyle.setFillForegroundColor(fillForegroundColor);
		}
	}

	@Deprecated
	public void setFillPattern(short fillPattern) {
		if (xssfCellStyle != null && fillPattern > 0) {
			this.fillPattern = fillPattern;
			xssfCellStyle.setFillPattern(fillPattern);
		}
	}

	public short getFillPattern() {
		return fillPattern;
	}

	// public short getFillForegroundColorColor() {
	// return fillForegroundColorColor;
	// }

	// public void setFillForegroundColorColor(short fillForegroundColorColor) {
	// if (xssfCellStyle != null) {
	// this.fillForegroundColorColor = fillForegroundColorColor;
	// xssfCellStyle.setFillForegroundColor(fill)
	// }
	// }

	public XSSFColor getFillForegroundXSSFColor() {
		return fillForegroundXSSFColor;
	}

	public void setFillForegroundXSSFColor(XSSFColor fillForegroundXSSFColor) {
		if (xssfCellStyle != null) {
			this.fillForegroundXSSFColor = fillForegroundXSSFColor;
			xssfCellStyle.setFillForegroundColor(fillForegroundXSSFColor);
		}
	}

	public FillPatternType getFillPatternEnum() {
		return fillPatternEnum;
	}

	public void setFillPatternEnum(FillPatternType fillPatternEnum) {
		if (xssfCellStyle != null) {
			this.fillPatternEnum = fillPatternEnum;
			xssfCellStyle.setFillPattern(fillPatternEnum);
		}
	}

	public XSSFFont getFont() {
		return font;
	}

	public void setFont(XSSFFont font) {
		if (xssfCellStyle != null) {
			this.font = font;
			xssfCellStyle.setFont(font);
		}
	}

	public short getFontIndex() {
		return fontIndex;
	}

	public void setFontIndex(short fontIndex) {
		if (xssfCellStyle != null && fontIndex > 0) {
			this.fontIndex = fontIndex;
			// xssfCellStyle.setfont
		}
	}

	public boolean isHidden() {
		return hidden;
	}

	public void setHidden(boolean hidden) {
		if (xssfCellStyle != null) {
			this.hidden = hidden;
			xssfCellStyle.setHidden(hidden);
		}
	}

	public short getIndention() {
		return indention;
	}

	public void setIndention(short indention) {
		if (xssfCellStyle != null && indention > 0) {
			this.indention = indention;
			xssfCellStyle.setIndention(indention);
		}
	}

	public short getIndex() {
		return index;
	}

	public void setIndex(short index) {
		if (xssfCellStyle != null && index > 0) {
			this.index = index;
			// xssfCellStyle.setinde
		}
	}

	public short getLeftBorderColor() {
		return leftBorderColor;
	}

	public void setLeftBorderColor(short leftBorderColor) {
		if (xssfCellStyle != null && leftBorderColor > 0) {
			this.leftBorderColor = leftBorderColor;
			xssfCellStyle.setLeftBorderColor(leftBorderColor);
		}
	}

	public XSSFColor getLeftBorderXSSFColor() {
		return leftBorderXSSFColor;
	}

	public void setLeftBorderXSSFColor(XSSFColor leftBorderXSSFColor) {
		if (xssfCellStyle != null) {
			this.leftBorderXSSFColor = leftBorderXSSFColor;
			xssfCellStyle.setLeftBorderColor(leftBorderXSSFColor);
		}
	}

	public boolean isLocked() {
		return locked;
	}

	public void setLocked(boolean locked) {
		if (xssfCellStyle != null) {
			this.locked = locked;
			xssfCellStyle.setLocked(locked);
		}
	}

	public short getRightBorderColor() {
		return rightBorderColor;
	}

	public void setRightBorderColor(short rightBorderColor) {
		if (xssfCellStyle != null && rightBorderColor > 0) {
			this.rightBorderColor = rightBorderColor;
			xssfCellStyle.setRightBorderColor(rightBorderColor);
		}
	}

	public XSSFColor getRightBorderXSSFColor() {
		return rightBorderXSSFColor;
	}

	public void setRightBorderXSSFColor(XSSFColor rightBorderXSSFColor) {
		if (xssfCellStyle != null) {
			this.rightBorderXSSFColor = rightBorderXSSFColor;
			xssfCellStyle.setRightBorderColor(rightBorderXSSFColor);
		}
	}

	public short getRotation() {
		return rotation;
	}

	public void setRotation(short rotation) {
		if (xssfCellStyle != null && rotation > 0) {
			this.rotation = rotation;
			xssfCellStyle.setRotation(rotation);
		}
	}

	public short getTopBorderColor() {
		return topBorderColor;
	}

	public void setTopBorderColor(short topBorderColor) {
		if (xssfCellStyle != null && topBorderColor > 0) {
			this.topBorderColor = topBorderColor;
			xssfCellStyle.setTopBorderColor(topBorderColor);
		}
	}

	public XSSFColor getTopBorderXSSFColor() {
		return topBorderXSSFColor;
	}

	public void setTopBorderXSSFColor(XSSFColor topBorderXSSFColor) {
		if (xssfCellStyle != null) {
			this.topBorderXSSFColor = topBorderXSSFColor;
			xssfCellStyle.setTopBorderColor(topBorderXSSFColor);
		}
	}

	public short getVerticalAlignment() {
		return verticalAlignment;
	}

	@Deprecated
	public void setVerticalAlignment(short verticalAlignment) {
		if (xssfCellStyle != null && verticalAlignment > 0) {
			this.verticalAlignment = verticalAlignment;
			xssfCellStyle.setVerticalAlignment(verticalAlignment);
		}
	}

	public VerticalAlignment getVerticalAlignmentEnum() {
		return verticalAlignmentEnum;
	}

	public void setVerticalAlignmentEnum(VerticalAlignment verticalAlignmentEnum) {
		if (xssfCellStyle != null) {
			this.verticalAlignmentEnum = verticalAlignmentEnum;
			xssfCellStyle.setVerticalAlignment(verticalAlignmentEnum);
		}
	}

	public boolean isWrapText() {
		return wrapText;
	}

	public void setWrapText(boolean wrapText) {
		if (xssfCellStyle != null) {
			this.wrapText = wrapText;
			xssfCellStyle.setWrapText(wrapText);
		}
	}

}
