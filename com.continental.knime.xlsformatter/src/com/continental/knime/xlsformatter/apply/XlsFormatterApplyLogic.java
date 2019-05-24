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

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.IdentityHashMap;
import java.util.stream.Collectors;

import org.apache.commons.lang3.time.DateUtils;
import org.apache.commons.lang3.tuple.Triple;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.ColorScaleFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold.RangeType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeLogger;

import com.continental.knime.xlsformatter.commons.AddressingTools;
import com.continental.knime.xlsformatter.commons.ColorTools;
import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.commons.XlsFormattingStateValidator;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellState;

/**
 * Holds the logic how to apply collected formatting instructions to an XLSX file via Apache POI.
 */
public class XlsFormatterApplyLogic {
  
	/**
	 * Number of different styles allowed in a workbook prior to execution. This is due to e.g. date/time columns
	 * being written by ExcelWriter with a corresponding defined style.
	 */
	private final static int ALLOWEDPREVIOUSSTYLES = 6;
	
	/**
	 * The threshold above which a successful execution warning will appear to nodes that load a XLS Formatter state
	 * with instructions yielding more required POI styles than the maximum * this quota.
	 */
	private final static double WARNINGTHRESHOLDSTYLEQUOTA = 0.8d;
	
	/**
	 * Applies an xls formatting instruction object to an xlsx file. 
	 * @param destinationFile The file path of the existing xlsx file to modify.
	 * @param sheetName The name of the sheet to modify or null to take the first sheet.
	 * @param xlsf The XLS Formatting instructions object.
	 * @param exec The execution context (for aborting the operation and providing progress information). 
	 * @throws IOException 
	 */
	public static void apply(
			final String inputFile,
			final String outputFile,
			final String sheetName,
			final XlsFormatterState xlsf,
			WarningMessageContainer warningMessageContainer,
			final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		// Open the file
		Workbook wb = null;
		CreationHelper createHelper = null;
		Sheet sheet = null;
		try (FileInputStream inputFileStream = new FileInputStream(inputFile)) {
			wb = WorkbookFactory.create(new FileInputStream(inputFile)); // returns XSSF workbook for XSLX files
			createHelper = wb.getCreationHelper();
		}
		catch (Exception e) {
			throw new IllegalArgumentException("Could not open XLS file (" + inputFile + "): " + e.toString() + ":" + e.getMessage(), e);
		}
				
		if (wb.getNumCellStyles() > ALLOWEDPREVIOUSSTYLES)
			throw new IllegalArgumentException("The input file already contains formatting styles. This is currently unsupported in our extension beyond a degree that KNIME's XLS Writer node would utilize.");
		
		
		// Find the desired sheet
		sheet = sheetName == null ? wb.getSheetAt(0) : wb.getSheet(sheetName);
		if (sheet == null)
			throw new Exception("Sheet " + (sheetName == null ? "0" : sheetName) + " not found in file " + inputFile);
		
		// Derive and generate necessary POI styles:
		Map<CellAddress, XSSFCellStyle> styleMap = deriveNecessaryStyles(wb, xlsf, exec, logger);
		
		// Loop all cells with instructions:
		Row row;
		org.apache.poi.ss.usermodel.Cell cell;
		boolean hasTypeConversionParsingErrors = false;
		boolean hasDataTypeInstructionOnNonStringCells = false;
		for (CellAddress cellAddress : xlsf.cells.keySet()) {
			exec.checkCanceled();
			
			// locate cell in POI model
			row = safelyGetRow(sheet, cellAddress.getRow());
			cell = safelyGetColumn(row, cellAddress.getColumn());
			
			// if the data type shall be converted from a String cell to something else, do it before setting the style
			CellState state = xlsf.cells.get(cellAddress);
			if (state != null && state.cellDataType != XlsFormatterState.CellDataType.UNMODIFIED) {
				if (cell.getCellTypeEnum() == CellType.STRING)
					switch (state.cellDataType) {
					case NUMERIC:
						double value;
						try {
							value = Double.parseDouble(cell.getStringCellValue());
							cell = row.createCell(cellAddress.getColumn());
							cell.setCellValue(value);
						} catch (NumberFormatException ne) {
							logger.warn("Could not parse numeric value \"" + cell.getStringCellValue() + "\" in cell " + cellAddress.toString());
							hasTypeConversionParsingErrors = true;
						}
						break;
					case BOOLEAN:
						String strValue = cell.getStringCellValue().trim();
						switch (strValue) {
						case "TRUE":
						case "1":
							cell = row.createCell(cellAddress.getColumn());
							cell.setCellValue(true);
							break;
						case "FALSE":
						case "0":
							cell = row.createCell(cellAddress.getColumn());
							cell.setCellValue(false);
							break;
						default:
							logger.warn("Could not parse boolean value \"" + cell.getStringCellValue() + "\" in cell " + cellAddress.toString());
							hasTypeConversionParsingErrors = true;
						}
						break;
					case LOCALDATE:
					case LOCALDATETIME:
					case LOCALTIME:
						Date date;
						try {
							date = DateUtils.parseDate(cell.getStringCellValue(), state.cellDataType.getDateTextFormats());
							cell = row.createCell(cellAddress.getColumn());
							cell.setCellValue(date);
						} catch (Exception e) {
							logger.warn("Could not parse date/time value \"" + cell.getStringCellValue() + "\" (expected format was \""  + state.cellDataType.getDateTextFormat() + "\") in cell " + cellAddress.toString());
							hasTypeConversionParsingErrors = true;
						}
						break;
					case FORMULA:
						throw new IllegalArgumentException("Setting formulas is not supported.");
					default:
						break;
					}
				else { // non string
					logger.warn("Could not change data type of cell " + cellAddress.toString() + " since it is not a String cell. Try changing the text format instead of a data type conversion.");
					hasDataTypeInstructionOnNonStringCells = true;
				}
			}
			
			// if there is a style to set, set it
			if (styleMap.containsKey(cellAddress))
				cell.setCellStyle(styleMap.get(cellAddress));
			
			// hyperlink
			String hyperlink = xlsf.cells.get(cellAddress).hyperlink;
			if (hyperlink != null) {
				XSSFHyperlink link = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
				link.setAddress(hyperlink);
				cell.setHyperlink((XSSFHyperlink)link);
			}
		}
		if (hasTypeConversionParsingErrors)
			warningMessageContainer.addMessage("Parsing error(s) during cell data type conversion. See log for details.");
		if (hasDataTypeInstructionOnNonStringCells)
			warningMessageContainer.addMessage("Data type conversion(s) on non-String cells could not be executed. See log for details.");
		
		// Fix window:
		if (xlsf.freezeSheetAtTopLeftCornerOfCell != null)
			sheet.createFreezePane(
					xlsf.freezeSheetAtTopLeftCornerOfCell.getColumn(),
					xlsf.freezeSheetAtTopLeftCornerOfCell.getRow());
		
		// Set column width:
		if (xlsf.columnWidths != null && xlsf.columnWidths.size() != 0)
			for (Integer c : xlsf.columnWidths.keySet()) {
				Double columnWidth = xlsf.columnWidths.get(c);
				if (columnWidth == null)
					sheet.autoSizeColumn(c);
				else
					sheet.setColumnWidth(c, xlsToPoiStandardColumnWidthConversion(columnWidth));
			}
		exec.checkCanceled();
		
		// Set row height:
		if (xlsf.rowHeights != null && xlsf.rowHeights.size() != 0)
			for (Integer r : xlsf.rowHeights.keySet())
				safelyGetRow(sheet, r).setHeightInPoints((float)(double)xlsf.rowHeights.get(r));
		exec.checkCanceled();
		
		// Hide columns:
		if (xlsf.hiddenColumns != null && xlsf.hiddenColumns.size() != 0)
			for (Integer c : xlsf.hiddenColumns)
				sheet.setColumnHidden(c, true);
		exec.checkCanceled();
		
		// Hide rows:
		if (xlsf.hiddenRows != null && xlsf.hiddenRows.size() != 0)
			for (Integer r : xlsf.hiddenRows)
				safelyGetRow(sheet, r).setZeroHeight(true);
		exec.checkCanceled();
		
		// auto-filter range
		if (xlsf.autoFilterRange != null)
			sheet.setAutoFilter(xlsf.autoFilterRange);
		
		// merge ranges
		if (xlsf.mergeRanges != null && xlsf.mergeRanges.size() != 0)
			for (CellRangeAddress range : xlsf.mergeRanges) {
				
				// to avoid invisible cells that re-appear after manually un-merging, delete everything but the range's top-left cell
				for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
					exec.checkCanceled();
					row = safelyGetRow(sheet, r);
					for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++)
						if (r != range.getFirstRow() || c != range.getFirstColumn()) {
							cell = row.getCell(c);
							if (cell != null)
								cell.setCellType(CellType.BLANK);
						}
				}
				
				sheet.addMergedRegion(range);
			}
		
		// conditional formatting:
		Map<String, List<CellAddress>> mapIdenticallyConditionalFormattedCells = new HashMap<String, List<CellAddress>>();
		for (CellAddress cellAddress : xlsf.cells.keySet()) {
			exec.checkCanceled();
			XlsFormatterState.ConditionalFormattingSet conditionalFormat = xlsf.cells.get(cellAddress).conditionalFormat; 
			if (conditionalFormat != null) {
				String key = conditionalFormat.toString();
				if (!mapIdenticallyConditionalFormattedCells.containsKey(key))
					mapIdenticallyConditionalFormattedCells.put(key, new ArrayList<CellAddress>());
				mapIdenticallyConditionalFormattedCells.get(key).add(cellAddress);
			}
		}
		if (mapIdenticallyConditionalFormattedCells.size() != 0) {
			SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
			for (String key : mapIdenticallyConditionalFormattedCells.keySet()) {
				XlsFormatterState.ConditionalFormattingSet conditionalFormat = xlsf.cells.get(
						mapIdenticallyConditionalFormattedCells.get(key).get(0)).conditionalFormat;
				if (conditionalFormat.backgroundScaleFixpoints.size() == 0)
					continue; // for now, only bgColor is implemented. Skip anything else
				List<CellRangeAddress> ranges = AddressingTools.getRangesFromAddressList(mapIdenticallyConditionalFormattedCells.get(key), exec, logger);
				
				ConditionalFormattingRule ruleColorScale = sheetCF.createConditionalFormattingColorScaleRule();
				ColorScaleFormatting cs1 = ruleColorScale.getColorScaleFormatting();
				Color[] bgColors = new Color[conditionalFormat.backgroundScaleFixpoints.size()];
				for (int x = 0; x < conditionalFormat.backgroundScaleFixpoints.size(); x++) {
					cs1.getThresholds()[x].setRangeType(RangeType.NUMBER);
					cs1.getThresholds()[x].setValue(conditionalFormat.backgroundScaleFixpoints.get(x).getLeft());
					bgColors[x] = ColorTools.getPoiColor(conditionalFormat.backgroundScaleFixpoints.get(x).getRight());
				}
				cs1.setColors(bgColors);
				CellRangeAddress[] rangesArray = new CellRangeAddress[ranges.size()];
				rangesArray = ranges.toArray(rangesArray);
				int poiCfId = sheetCF.addConditionalFormatting(rangesArray, ruleColorScale);
				logger.debug("Setting conditional formatting (POI ID " + poiCfId + ") \"" + key + "\" for " + ranges.size() + " ranges: " + ranges.stream().map(r -> r.formatAsString()).collect(Collectors.joining(";")));
				int cfAddressFormulaLength = AddressingTools.getFormulaLengthFromRangeList(ranges);
				if (cfAddressFormulaLength > XlsFormattingStateValidator.MAX_FORMULA_CHARACTERS) {
					logger.warn("This conditional formatting address range's length exceeds the XLS specification's maximum formula character limit: " + cfAddressFormulaLength + " > " + XlsFormattingStateValidator.MAX_FORMULA_CHARACTERS);
					warningMessageContainer.addMessage("A conditional formatting range is so jagged that it exceeds the maximum XLS formula length. See log for details.");
				}
			}
		}
		exec.checkCanceled();
		
		// Group columns:
		if (xlsf.columnGroups != null && xlsf.columnGroups.size() != 0)
			for (Triple<Integer, Integer, Boolean> group : xlsf.columnGroups) { // from, to, collapsed
				sheet.groupColumn(group.getLeft(), group.getMiddle());
				sheet.setColumnGroupCollapsed(group.getLeft(), group.getRight());
			}
		exec.checkCanceled();
		
		// Group rows:
		if (xlsf.rowGroups != null && xlsf.rowGroups.size() != 0)
			for (Triple<Integer, Integer, Boolean> group : xlsf.rowGroups) { // from, to, collapsed
				sheet.groupRow(group.getLeft(), group.getMiddle());
				sheet.setRowGroupCollapsed(group.getLeft(), group.getRight());
			}
		exec.checkCanceled();
		
		// Write the output to a file:
		try (FileOutputStream fileOut = new FileOutputStream(outputFile)) {
			wb.write(fileOut);
		}
  }
	
	
	/**
	 * Safely accesses a row on a POI sheet by creating it in case it doesn't exist.
	 * @param sheet
	 * @param rowIndex 0-based row index.
	 */
	private static Row safelyGetRow(Sheet sheet, int rowIndex) {
		Row row = sheet.getRow(rowIndex);
		if (row == null)
			row = sheet.createRow(rowIndex);
		return row;
	}
	
	/**
	 * Safely accesses a cell in a POI row by creating it in case it doesn't exist.
	 * @param row The POI row object.
	 * @param colIndex 0-based column index.
	 */
	private static Cell safelyGetColumn(Row row, int colIndex) {
		Cell cell = row.getCell(colIndex);
		if (cell == null)
			cell = row.createCell(colIndex);
		return cell;
	}
	
	
	/**
	 * Analyze the XLS Formatting instructions for common styles and generate these styles.
	 * @return A map that contains the to define style for every cell address that needs to be formatted.
	 */
	private static Map<CellAddress, XSSFCellStyle> deriveNecessaryStyles(
			final Workbook workbook,
			final XlsFormatterState xlsf,
			final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		Map<CellAddress, XSSFCellStyle> ret = new HashMap<CellAddress, XSSFCellStyle>();
		
		Map<String, XSSFFont> fontMap = new HashMap<String, XSSFFont>(); // maps font short code to created POI font
		Map<String, XSSFCellStyle> styleMap = new HashMap<String, XSSFCellStyle>(); // maps cell style short code to created POI cell style
		Map<String, Integer> numberFormatMap = new HashMap<String, Integer>(); // maps number format string to created POI format code
		
		for (CellAddress cellAddress : xlsf.cells.keySet()) {
			try {
				XlsFormatterState.CellState state = xlsf.cells.get(cellAddress);
								
				// check whether the cell state requires any formatting:
				String cellShortString = state.cellFormatToShortString(true); 
				if (cellShortString.equals(XlsFormatterState.CellState.getNonFormattingStateString()))
					continue;
				
				// resolve font:
				String fontShortString = state.fontDefinitionToShortString();
				XSSFFont font = null;
				if (!fontShortString.equals(XlsFormatterState.CellState.getDefaultFontShortString())) {
					if (fontMap.containsKey(fontShortString))
						font = fontMap.get(fontShortString);
					else {
						font = XlsFormatterApplyXlsfToPoiConversions.createFont(workbook, state);
						fontMap.put(fontShortString, font);
					}
				}
				
				// resolve number format:
				String targetNumberFormat = state.textFormat;
				if (targetNumberFormat == null) // in case the user didn't define the format, but the data type conversion requires a format, define it here
					switch (state.cellDataType) {
					case LOCALDATE:
					case LOCALDATETIME:
					case LOCALTIME:
						targetNumberFormat = state.cellDataType.getDateTextFormat();
						break;
					default:
						break;
					}				
				Integer numberFormatCode = null;
				if (targetNumberFormat != null) {
					if (numberFormatMap.containsKey(targetNumberFormat))
						numberFormatCode = numberFormatMap.get(targetNumberFormat);
					else {
						numberFormatCode = workbook == null ? -1 : (int)workbook.createDataFormat().getFormat(targetNumberFormat);
						numberFormatMap.put(targetNumberFormat, numberFormatCode);
					}
				}
				
				// resolve style:
				XSSFCellStyle style = null;
				if (styleMap.containsKey(cellShortString))
					style = styleMap.get(cellShortString);
				else {
					style = XlsFormatterApplyXlsfToPoiConversions.createCellStyle(workbook, state, font, numberFormatCode);
					styleMap.put(cellShortString, style);
				}
				
				ret.put(cellAddress, style);
			}
			catch (Exception e) {
				StringWriter sw = new StringWriter();
				PrintWriter pw = new PrintWriter(sw);
				e.printStackTrace(pw);
				logger.error(e.getClass().getCanonicalName() + ": " + e.getMessage() + "\n" + sw.toString());
				throw new Exception(cellAddress.formatAsString() + ": " + e.getClass().getCanonicalName() + ":" + e.getMessage(), e);
			}
		}
		return ret;
	}
	
	/**
	 * Safely get the simulated number of POI styles that would be created to implement the provided XlsFormatterState.
	 * @param xlsf
	 * @param exec
	 * @param logger
	 * @return
	 * @throws Exception
	 */
	public static int getNumberOfNecessaryStyles(final XlsFormatterState xlsf,
			final ExecutionContext exec, final NodeLogger logger) throws Exception {
		Set<XSSFCellStyle> set = Collections.newSetFromMap(new IdentityHashMap<>()); // workaround to achieve a duplicate resolution based on object identity instead of content
		set.addAll(deriveNecessaryStyles(null, xlsf, exec, logger).values()); // deriveNecessaryStyles called without a workbook creates a dummy XSSFCellStyle per new Style String (but all of these are content wise identical, hence their total count is determined via the identity-based set workaround of the previous line. 
		return set.size();
	}
	
	public static void checkDerivedStyleComplexity(XlsFormatterState state, WarningMessageContainer warningMessageContainer,
			final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		int requiredStyles = getNumberOfNecessaryStyles(state, exec, logger);
		
		if (requiredStyles > XlsFormattingStateValidator.MAX_CELL_STYLES_PER_WORKBOOK - ALLOWEDPREVIOUSSTYLES)
			throw new Exception("The XLS Formatter port object has been loaded with more instructions than can be implemented in an XLS workbook (" + XlsFormattingStateValidator.MAX_CELL_STYLES_PER_WORKBOOK + ").");
		
		double requiredStyleQuotaUsage = requiredStyles / (double)XlsFormattingStateValidator.MAX_CELL_STYLES_PER_WORKBOOK;
		if (requiredStyleQuotaUsage >	WARNINGTHRESHOLDSTYLEQUOTA)
			warningMessageContainer.addMessage("The XLS Formatter port object has already been loaded with " + Math.round(requiredStyleQuotaUsage * 100) + "% of the allowed maximum.");
		
		logger.debug("XLS Formatter port object validation: " + requiredStyles + " of " + XlsFormattingStateValidator.MAX_CELL_STYLES_PER_WORKBOOK + " (" + Math.round(requiredStyleQuotaUsage * 100) + "%) allowed styles.");
	}
	
	/**
	 * Converts the column width shown in a standard spreadsheet to that required by POI.
	 * Depends on some font setting, we assume standard here (written by POI, read by POI).
	 */
	private static int xlsToPoiStandardColumnWidthConversion(double xlsColumnWidth) {
		return (int)Math.round(257.5 * xlsColumnWidth + 165.37); // formula based on own experiment with different settings and linear regression on that
	}
}
