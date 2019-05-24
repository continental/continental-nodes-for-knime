/*
 * Continental Nodes for KNIME
 * Copyright (C) 2019  Continental AG, Hanover, Germany
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <https://www.gnu.org/licenses/>.
 */

package com.continental.knime.xlsformatter.apply;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;

import com.continental.knime.xlsformatter.commons.ColorTools;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellAlignmentHorizontal;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellAlignmentVertical;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.FillPattern;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.FormattingFlag;


/**
 * Holds the logic how to apply collected formatting instructions to an XLSX file via Apache POI.
 */
public class XlsFormatterApplyXlsfToPoiConversions {
    
	/**
	 * Creates a POI font based on an XLS Formatting cell state.
	 */
	public static XSSFFont createFont(Workbook workbook, CellState cellState) {
		
		if (workbook == null)
			return null;
		
		XSSFFont font = ((XSSFWorkbook)workbook).createFont();
		
		if (cellState.fontSize != null)
			font.setFontHeightInPoints((short)(int)cellState.fontSize);
		
		if (cellState.fontBold != FormattingFlag.UNMODIFIED)
			font.setBold(cellState.fontBold == FormattingFlag.ON);
    
		if (cellState.fontUnderline != FormattingFlag.UNMODIFIED)
			font.setUnderline(cellState.fontUnderline == FormattingFlag.ON ? XSSFFont.U_SINGLE : XSSFFont.U_NONE);
		
		if (cellState.fontItalic != FormattingFlag.UNMODIFIED)
			font.setItalic(cellState.fontItalic == FormattingFlag.ON);
		
		if (cellState.fontColor != null)
			font.setColor(ColorTools.getPoiColor(cellState.fontColor));
		
		return font;
	}
	
	/**
	 * Creates a POI cell style, based on XLS Formatting cell state, a pre-generated POI font (can be null), and a pre-generated POI number format (can be null).
	 */
	public static XSSFCellStyle createCellStyle(Workbook workbook, CellState cellState, XSSFFont font, Integer numberFormatCode) {
		
		if (workbook == null)
			return new XSSFCellStyle(null);
		
		XSSFCellStyle style = (XSSFCellStyle)workbook.createCellStyle();
		
		XSSFColor poiBackgroundColor = null;
		
		if (cellState.backgroundColor != null) {
			poiBackgroundColor = ColorTools.getPoiColor(cellState.backgroundColor);
			style.setFillBackgroundColor(poiBackgroundColor);
		}
		
		if (cellState.backgroundColor != null && cellState.fillPattern == FillPattern.SOLID_BACKGROUND_COLOR) {
			style.setFillForegroundColor(poiBackgroundColor);
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		} else if (cellState.fillForegroundColor != null)
			style.setFillForegroundColor(ColorTools.getPoiColor(cellState.fillForegroundColor));
		
		if (cellState.fillPattern == FillPattern.NONE)
			style.setFillPattern(FillPatternType.NO_FILL);
		else if (cellState.fillPattern != FillPattern.UNMODIFIED && cellState.fillPattern != FillPattern.SOLID_BACKGROUND_COLOR)
			style.setFillPattern(getFillPatternXlsfToPoi(cellState.fillPattern));
		
		
		if (cellState.cellHorizontalAlignment != CellAlignmentHorizontal.UNMODIFIED)
			style.setAlignment(getHorizontalAlignmentXlsfToPoi(cellState.cellHorizontalAlignment));
		
		if (cellState.cellVerticalAlignment != CellAlignmentVertical.UNMODIFIED)
			style.setVerticalAlignment(getVerticalAlignmentXlsfToPoi(cellState.cellVerticalAlignment));
		
		if (cellState.wrapText != FormattingFlag.UNMODIFIED)
			style.setWrapText(cellState.wrapText == FormattingFlag.ON);
		
		if (cellState.textTiltDegree != null)
			style.setRotation((short)(int)cellState.textTiltDegree);
		
		if (cellState.borderTop != null && cellState.borderTop.style != XlsFormatterState.BorderStyle.UNMODIFIED)
			style.setBorderTop(getBorderStyleXlsfToPoi(cellState.borderTop.style));
		if (cellState.borderTop != null && cellState.borderTop.color != null)
			style.setBorderColor(BorderSide.TOP, ColorTools.getPoiColor(cellState.borderTop.color));
		
		if (cellState.borderLeft != null && cellState.borderLeft.style != XlsFormatterState.BorderStyle.UNMODIFIED)
			style.setBorderLeft(getBorderStyleXlsfToPoi(cellState.borderLeft.style));
		if (cellState.borderLeft != null && cellState.borderLeft.color != null)
			style.setBorderColor(BorderSide.LEFT, ColorTools.getPoiColor(cellState.borderLeft.color));
		
		if (cellState.borderBottom != null && cellState.borderBottom.style != XlsFormatterState.BorderStyle.UNMODIFIED)
			style.setBorderBottom(getBorderStyleXlsfToPoi(cellState.borderBottom.style));
		if (cellState.borderBottom != null && cellState.borderBottom.color != null)
			style.setBorderColor(BorderSide.BOTTOM, ColorTools.getPoiColor(cellState.borderBottom.color));
		
		if (cellState.borderRight != null && cellState.borderRight.style != XlsFormatterState.BorderStyle.UNMODIFIED)
			style.setBorderRight(getBorderStyleXlsfToPoi(cellState.borderRight.style));
		if (cellState.borderRight != null && cellState.borderRight.color != null)
			style.setBorderColor(BorderSide.RIGHT, ColorTools.getPoiColor(cellState.borderRight.color));
		
		if (numberFormatCode != null)
			style.setDataFormat(numberFormatCode);		
		
		if (font != null)
			style.setFont(font);
		
		return style;
	}
	
	
	/*** Maps of XLSF to POI enums ***/
	private static Map<CellAlignmentHorizontal, HorizontalAlignment> _mapHorizontalAlignmentXlsfToPoi = null;
	private static HorizontalAlignment getHorizontalAlignmentXlsfToPoi(CellAlignmentHorizontal value) {
		if (_mapHorizontalAlignmentXlsfToPoi == null) {
			_mapHorizontalAlignmentXlsfToPoi = new HashMap<CellAlignmentHorizontal, HorizontalAlignment>();
			_mapHorizontalAlignmentXlsfToPoi.put(CellAlignmentHorizontal.LEFT, HorizontalAlignment.LEFT);
			_mapHorizontalAlignmentXlsfToPoi.put(CellAlignmentHorizontal.RIGHT, HorizontalAlignment.RIGHT);
			_mapHorizontalAlignmentXlsfToPoi.put(CellAlignmentHorizontal.CENTER, HorizontalAlignment.CENTER);
			_mapHorizontalAlignmentXlsfToPoi.put(CellAlignmentHorizontal.JUSTIFY, HorizontalAlignment.JUSTIFY);
		}
		return _mapHorizontalAlignmentXlsfToPoi.get(value);
	}
	
	private static Map<CellAlignmentVertical, VerticalAlignment> _mapVerticalAlignmentXlsfToPoi = null;
	private static VerticalAlignment getVerticalAlignmentXlsfToPoi(CellAlignmentVertical value) {
		if (_mapVerticalAlignmentXlsfToPoi == null) {
			_mapVerticalAlignmentXlsfToPoi = new HashMap<CellAlignmentVertical, VerticalAlignment>();
			_mapVerticalAlignmentXlsfToPoi.put(CellAlignmentVertical.TOP, VerticalAlignment.TOP);
			_mapVerticalAlignmentXlsfToPoi.put(CellAlignmentVertical.MIDDLE, VerticalAlignment.CENTER);
			_mapVerticalAlignmentXlsfToPoi.put(CellAlignmentVertical.BOTTOM, VerticalAlignment.BOTTOM);
		}
		return _mapVerticalAlignmentXlsfToPoi.get(value);
	}
	
	private static Map<XlsFormatterState.BorderStyle, org.apache.poi.ss.usermodel.BorderStyle> _mapBorderStyleXlsfToPoi = null;
	private static org.apache.poi.ss.usermodel.BorderStyle getBorderStyleXlsfToPoi(XlsFormatterState.BorderStyle value) {
		if (_mapBorderStyleXlsfToPoi == null) {
			_mapBorderStyleXlsfToPoi = new HashMap<XlsFormatterState.BorderStyle, org.apache.poi.ss.usermodel.BorderStyle>();
			_mapBorderStyleXlsfToPoi.put(XlsFormatterState.BorderStyle.NORMAL, org.apache.poi.ss.usermodel.BorderStyle.THIN);
			_mapBorderStyleXlsfToPoi.put(XlsFormatterState.BorderStyle.THICK, org.apache.poi.ss.usermodel.BorderStyle.MEDIUM);
			_mapBorderStyleXlsfToPoi.put(XlsFormatterState.BorderStyle.EXTRA_THICK, org.apache.poi.ss.usermodel.BorderStyle.THICK);
			_mapBorderStyleXlsfToPoi.put(XlsFormatterState.BorderStyle.DOUBLE, org.apache.poi.ss.usermodel.BorderStyle.DOUBLE);
			_mapBorderStyleXlsfToPoi.put(XlsFormatterState.BorderStyle.DASHED, org.apache.poi.ss.usermodel.BorderStyle.DASHED);
			_mapBorderStyleXlsfToPoi.put(XlsFormatterState.BorderStyle.DASHED_THICK, org.apache.poi.ss.usermodel.BorderStyle.MEDIUM_DASHED);
			_mapBorderStyleXlsfToPoi.put(XlsFormatterState.BorderStyle.NONE, org.apache.poi.ss.usermodel.BorderStyle.NONE);
		}
		return _mapBorderStyleXlsfToPoi.get(value);
	}
	
	private static Map<FillPattern, FillPatternType> _mapFillPatternXlsfToPoi = null;
	private static FillPatternType getFillPatternXlsfToPoi(FillPattern value) {
		if (_mapFillPatternXlsfToPoi == null) {
			_mapFillPatternXlsfToPoi = new HashMap<FillPattern, FillPatternType>();
			_mapFillPatternXlsfToPoi.put(FillPattern.DIAGONAL, FillPatternType.THIN_FORWARD_DIAG);
			_mapFillPatternXlsfToPoi.put(FillPattern.HORIZONTAL, FillPatternType.THIN_HORZ_BANDS);
			_mapFillPatternXlsfToPoi.put(FillPattern.VERTICAL, FillPatternType.THIN_VERT_BANDS);
			_mapFillPatternXlsfToPoi.put(FillPattern.DOTTED, FillPatternType.LESS_DOTS);
			// unmodified, none, and solid_background_color shall never be looked up here but caught in the calling logic
		}
		return _mapFillPatternXlsfToPoi.get(value);
	}	
}
