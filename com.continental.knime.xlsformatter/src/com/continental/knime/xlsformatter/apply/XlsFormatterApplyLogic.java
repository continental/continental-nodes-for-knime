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

import java.io.BufferedOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.commons.lang3.time.DateUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.ColorScaleFormatting;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold.RangeType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.RichTextString;
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
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeLogger;
import org.knime.filehandling.core.util.CheckedExceptionSupplier;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheetViews;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

import com.continental.knime.xlsformatter.commons.AddressingTools;
import com.continental.knime.xlsformatter.commons.ColorTools;
import com.continental.knime.xlsformatter.commons.ProgressTools;
import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.commons.XlsFormattingStateValidator;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.SheetState;

/**
 * Holds the logic how to apply collected formatting instructions to an XLSX file via Apache POI.
 */
public class XlsFormatterApplyLogic {
  
	/**
	 * Number of different styles allowed in a workbook prior to execution. This is due to e.g. date/time columns
	 * being written by ExcelWriter with a corresponding defined style.
	 */
	private final static int ALLOWED_PREVIOUS_STYLES = 6;
	
	/**
	 * The threshold above which a successful execution warning will appear to nodes that load a XLS Formatter state
	 * with instructions yielding more required POI styles than the maximum * this quota.
	 */
	private final static double WARNING_THRESHOLD_QUOTA = 0.8d;
	
	/**
	 * An object that is returned by the derive necessary styles logic to count the number of XLS / POI artifacts
	 * necessary to implement the desired XLS Formatting instruction, but that the XLS specification has a quota on.
	 */
	public static class XlsArtifactCount {
		public int StyleCount = 1;
		public int FontCount = 1;
		public int NumberFormatCount = 1;
	}
	
	/**
	 * Applies an XLS Formatting instruction object to an xlsx file.
	 * 
	 * @param inputFile  the input file.
	 * @param outputFile The file path of the existing xlsx file to modify.
	 * @param xlsf       The XLS Formatting instructions object.
	 * @param exec       The execution context (for aborting the operation and
	 *                   providing progress information).
	 * @param logger     the node logger.
	 * @throws IOException
	 */
	public static void apply(final String inputFile, final String outputFile, final XlsFormatterState xlsf,
			WarningMessageContainer warningMessageContainer, final ExecutionContext exec, final NodeLogger logger)
			throws Exception {
		apply(inputFile, //
				() -> new FileInputStream(inputFile), //
				() -> new FileOutputStream(outputFile), //
				xlsf, //
				warningMessageContainer, //
				exec, //
				logger);
	}

	/**
	 * Applies an XLS Formatting instruction object to an xlsx file.
	 * 
	 * @param inputFile  the input file name.
	 * @param openInput  creates the input stream to read from.
	 * @param openOutput create the output stream to write to.
	 * @param xlsf       The XLS Formatting instructions object.
	 * @param exec       The execution context (for aborting the operation and
	 *                   providing progress information).
	 * @param logger     the node logger.
	 * @throws IOException
	 */
	public static void apply(
			final String inputFile,
			final CheckedExceptionSupplier<InputStream, IOException> openInput,
			final CheckedExceptionSupplier<OutputStream, IOException> openOutput,
			final XlsFormatterState xlsf,
			WarningMessageContainer warningMessageContainer,
			final ExecutionContext exec, final NodeLogger logger) throws Exception {
		// Open the file
		Workbook wb = null;
		CreationHelper createHelper = null;
		exec.setProgress("Opening input file...");
		try (InputStream inputFileStream = openInput.get()) {
			wb = WorkbookFactory.create(inputFileStream); // returns XSSF workbook for XSLX files
			createHelper = wb.getCreationHelper();
		}
		catch (Exception e) {
			throw new IllegalArgumentException("Could not open XLS file (" + inputFile + "): " + e.toString() + ":" + e.getMessage(), e);
		}
		
		int numberOfPreviousStyles = wb.getNumCellStyles();
		logger.debug(numberOfPreviousStyles + " previous style(s) and " + wb.getNumberOfFontsAsInt() + " font(s) found in input file " + inputFile); 
		
		if (numberOfPreviousStyles > ALLOWED_PREVIOUS_STYLES) {
			
			// Analyze these pre-existing styles (esp. since KNIME's Sheet Appender node seems to add duplicate styles):
			int numberOfUniquePreviousStyles = 0;
			for (int i = 0; i < numberOfPreviousStyles; i++) {
				boolean isDuplicate = false;
				for (int j = 0; j < i && !isDuplicate; j++)
					if (wb.getCellStyleAt(i).equals(wb.getCellStyleAt(j)))
						isDuplicate = true;
				if (!isDuplicate)
					numberOfUniquePreviousStyles++;
			}
			logger.debug("  Thereof " + numberOfUniquePreviousStyles + " unique style(s). Our tolerance limit is " + ALLOWED_PREVIOUS_STYLES + ".");
			
			if (numberOfUniquePreviousStyles > ALLOWED_PREVIOUS_STYLES)
				throw new IllegalArgumentException("The input file already contains formatting styles. This is currently unsupported in our extension beyond a degree that KNIME's XLS Writer and Sheet Appender nodes would utilize.");
		}
		
		// Note that if we carry on with some pre-existing (but duplicate) styles, adding many new styles could exceed the total style limit. This might not be detected
		// in deriveNecessaryStyles yet, but would still trigger a POI-caused exception in the (very time consuming) process after adding the n+1st style
		
		// Derive and generate necessary POI styles:
		exec.setProgress("Adding necessary styles...");
		StyleAnalysisResult analysisResult = deriveNecessaryStyles(wb, xlsf, exec, logger);
		
		// Loop all sheets
		exec.setProgress("Applying formatting instructions...");
		Sheet defaultSheetIfAddressed = xlsf.sheetStates.containsKey(null) ? wb.getSheetAt(0) : null;
		for (String sheetName : xlsf.sheetStates.keySet()) {
			
			// Find the desired sheet
			Sheet sheet = sheetName == null ? wb.getSheetAt(0) : wb.getSheet(sheetName);
			if (sheet == null)
				throw new Exception("Sheet " + (sheetName == null ? "0" : "\"" + sheetName + "\"") + " not found in file " + inputFile);
			if (sheetName != null && sheet == defaultSheetIfAddressed)
				throw new Exception("Default sheet (i.e. first in sequence) and named sheet \"" + sheetName + "\" have been addressed separately. This is not supported as potentially conflicting formatting instructions could be written to the very same sheet.");
			XlsFormatterState.SheetState xlsfs = xlsf.sheetStates.get(sheetName);
			
			// prepare POI workbook level objects for adding drawing (i.e. for cell comments), which will only be instantiated jit:
			Drawing<?> drawing = null;
	    ClientAnchor clientAnchor = null;
			
			// Loop all cells with instructions:
			Row row;
			org.apache.poi.ss.usermodel.Cell cell;
			boolean hasTypeConversionParsingErrors = false;
			boolean hasDataTypeInstructionOnNonStringCells = false;
			int cellCount = xlsfs.cells.size();
			int cellIterator = 1;
			for (CellAddress cellAddress : xlsfs.cells.keySet()) {
				exec.checkCanceled();
				ProgressTools.showProgressText(exec, "Applying formatting instructions to cell", cellIterator++, cellCount, "...");
				
				// locate cell in POI model
				row = safelyGetRow(sheet, cellAddress.getRow());
				cell = safelyGetColumn(row, cellAddress.getColumn());
				
				// if the data type shall be converted from a String cell to something else, do it before setting the style
				CellState state = xlsfs.cells.get(cellAddress);
				if (state != null && state.cellDataType != XlsFormatterState.CellDataType.UNMODIFIED) {
					if (cell.getCellType() == CellType.STRING)
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
				Pair<String, CellAddress> sheetCellAddress = Pair.of(sheetName, cellAddress);
				if (analysisResult.mapCellAddressToStyleId.containsKey(sheetCellAddress))
					cell.setCellStyle(
							analysisResult.mapStyleIdToPoiStyle.get(
									analysisResult.mapCellAddressToStyleId.get(sheetCellAddress)));
				
				// hyperlink
				String hyperlink = state.hyperlink;
				if (hyperlink != null) {
					XSSFHyperlink link = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
					link.setAddress(hyperlink);
					cell.setHyperlink((XSSFHyperlink)link);
				}
				
				// comment
				if (state.comment != null && state.comment.text != null) {
					if (drawing == null)
						drawing = sheet.createDrawingPatriarch();
					if (clientAnchor == null)
						clientAnchor = createHelper.createClientAnchor();
					
					 org.apache.poi.ss.usermodel.Comment comment = drawing.createCellComment(clientAnchor);
           RichTextString richString = createHelper.createRichTextString(state.comment.text);
           comment.setString(richString);
           comment.setAuthor(state.comment.author); // can be null
           cell.setCellComment(comment);
				}
			} // for each cell
			
			// show warnings: 
			if (hasTypeConversionParsingErrors)
				warningMessageContainer.addMessage("Parsing error(s) during cell data type conversion. See log for details.");
			if (hasDataTypeInstructionOnNonStringCells)
				warningMessageContainer.addMessage("Data type conversion(s) on non-String cells could not be executed. See log for details.");
			
			// Fix window:
			exec.setProgress("Apply non-cell based features...");
			if (xlsfs.freezeSheetAtTopLeftCornerOfCell != null) {
				boolean errorInRepositioningView = false;
				if (sheet.getTopRow() != 0 || sheet.getLeftCol() != 0)  { // set the view to beginning of sheet, since the below createFreezePane freezes that view
					logger.debug("Sheet view repositioning necessary before sheet can be frozen since the input file was presumably saved at a window scroll position not displaying the top left of the sheet.");
					try {
						CTWorksheet ctSheet = ((XSSFSheet)sheet).getCTWorksheet();
						CTSheetViews ctSheetViews = ctSheet.getSheetViews(); 
						ctSheetViews.getSheetViewArray(ctSheetViews.sizeOfSheetViewArray() - 1).setTopLeftCell("A1");
					} catch (Exception e) {
						warningMessageContainer.addMessage("The input file was saved at a window scroll position not displaying the top left of the sheet. Re-positioning failed, hence the sheet freeze instruction needed to be skipped.");
						errorInRepositioningView = true;
					}
				}
				if (!errorInRepositioningView)
					sheet.createFreezePane(
							xlsfs.freezeSheetAtTopLeftCornerOfCell.getColumn(),
							xlsfs.freezeSheetAtTopLeftCornerOfCell.getRow());
			}
			
			// merge ranges
			if (xlsfs.mergeRanges != null && xlsfs.mergeRanges.size() != 0)
				for (CellRangeAddress range : xlsfs.mergeRanges) {
					
					// to avoid invisible cells that re-appear after manually un-merging, delete everything but the range's top-left cell
					for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
						exec.checkCanceled();
						row = safelyGetRow(sheet, r);
						for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++)
							if (r != range.getFirstRow() || c != range.getFirstColumn()) {
								cell = row.getCell(c);
								if (cell != null) {
									cell.setBlank();
									cell.removeCellComment();
									cell.removeHyperlink();
								}
							}
					}
					
					sheet.addMergedRegion(range);
				}
			
			// auto-filter range
			if (xlsfs.autoFilterRange != null)
				sheet.setAutoFilter(xlsfs.autoFilterRange);
			
			// conditional formatting:
			Map<String, List<CellAddress>> mapIdenticallyConditionalFormattedCells = new HashMap<String, List<CellAddress>>();
			for (CellAddress cellAddress : xlsfs.cells.keySet()) {
				exec.checkCanceled();
				XlsFormatterState.ConditionalFormattingSet conditionalFormat = xlsfs.cells.get(cellAddress).conditionalFormat; 
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
					XlsFormatterState.ConditionalFormattingSet conditionalFormat = xlsfs.cells.get(
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
			if (xlsfs.columnGroups != null && xlsfs.columnGroups.size() != 0)
				for (Map.Entry<Pair<Integer, Integer>, Boolean> group : xlsfs.columnGroups.entrySet()) { // <from, to>, collapsed
					sheet.groupColumn(group.getKey().getLeft(), group.getKey().getRight());
					sheet.setColumnGroupCollapsed(group.getKey().getLeft(), group.getValue()); // here, the group is identified by its from id
				}
			exec.checkCanceled();
			
			// Group rows:
			if (xlsfs.rowGroups != null && xlsfs.rowGroups.size() != 0)
				for (Map.Entry<Pair<Integer, Integer>, Boolean> group : xlsfs.rowGroups.entrySet()) { // <from, to>, collapsed
					sheet.groupRow(group.getKey().getLeft(), group.getKey().getRight());
					sheet.setRowGroupCollapsed(group.getKey().getLeft(), group.getValue()); // here, the group is identified by its from id
				}
			exec.checkCanceled();
			
			// Hide columns:
			if (xlsfs.hiddenColumns != null && xlsfs.hiddenColumns.size() != 0)
				for (Integer c : xlsfs.hiddenColumns)
					sheet.setColumnHidden(c, true);
			exec.checkCanceled();
			
			// Hide rows:
			if (xlsfs.hiddenRows != null && xlsfs.hiddenRows.size() != 0)
				for (Integer r : xlsfs.hiddenRows)
					safelyGetRow(sheet, r).setZeroHeight(true);
			exec.checkCanceled();
			
			// Set row height:
			if (xlsfs.rowHeights != null && xlsfs.rowHeights.size() != 0)
				for (Integer r : xlsfs.rowHeights.keySet())
					safelyGetRow(sheet, r).setHeightInPoints((float)(double)xlsfs.rowHeights.get(r));
			exec.checkCanceled();
			
			// Set column width (note that this should come last as auto-size can be dependent on other formatting settings):
			if (xlsfs.columnWidths != null && xlsfs.columnWidths.size() != 0)
				for (Integer c : xlsfs.columnWidths.keySet()) {
					Double columnWidth = xlsfs.columnWidths.get(c);
					if (columnWidth == null)
						sheet.autoSizeColumn(c);
					else
						sheet.setColumnWidth(c, xlsToPoiStandardColumnWidthConversion(columnWidth));
				}
			exec.checkCanceled();
		}
		
		// Write the output to a file:
		exec.setProgress("Writing output file...");
		try (OutputStream fileOut = openOutput.get();
				BufferedOutputStream bufOut = new BufferedOutputStream(fileOut);) {
			wb.write(bufOut);
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
	 * An object that is used to store style creation instruction, in order to
	 * create all necessary styles in one pass (after analyzing the required styles)
	 */
	public static class StyleCreationInstruction {
		public CellState cellState;
		public XSSFFont font;
		public Integer numberFormatCode; // null means: nothing to tamper with in regards to number formats
		
		public StyleCreationInstruction(CellState cellState, XSSFFont font, Integer numberFormatCode) {
			this.cellState = cellState;
			this.numberFormatCode = numberFormatCode;
			this.font = font;
		}
	}
	
	/**
	 * Result of the style derivation algorithm.
	 */
	public static class StyleAnalysisResult {
		
		/**
		 * Statistics of required artifacts such as styles, fonts, and number formats.
		 */
		XlsArtifactCount xlsArtifactCount = null;
		
		/**
		 * Map connecting a pair of sheet name & cell address to an internal counting ID connecting this map and mapStyleCodeToPoiStyle
		 */
		Map<Pair<String, CellAddress>, Integer> mapCellAddressToStyleId = null;
		
		/**
		 * Map connecting the internal counting ID of mapCellAddressToStyleCode to the corresponding POI XSSF style
		 */
		Map<Integer, XSSFCellStyle> mapStyleIdToPoiStyle = null;
	}
	
	/**
	 * Analyze the XLS Formatting instructions for common styles and generate these styles.
	 * @param workbook The POI workbook to create the styles in or null if the intention is only to calculate necessary artifact counts.
	 * @param xlsArtifactCount A prepared artifact counter that will be modified to reflect the XLS artifact utilization required to implement this XLS Formatting instruction set. Can be null, if this additional return is of no interest.
	 * @return A StyleAnalysisResult object (with only xlsArtifactCount filled in the case of workbook being null).
	 */
	private static StyleAnalysisResult deriveNecessaryStyles(
			final Workbook workbook,
			final XlsFormatterState xlsf,
			final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		StyleAnalysisResult ret = new StyleAnalysisResult();
		
		Map<String, XSSFFont> fontMap = new HashMap<String, XSSFFont>(); // maps font short code to created POI font
		Map<Integer, StyleCreationInstruction> styleIdToCreationInstructionMap = new HashMap<Integer, StyleCreationInstruction>(); // maps style internal ID to its creation instruction (for later bulk creation of styles)
		Map<String, Integer> styleCodeToInternalIdMap = new HashMap<String, Integer>(); // maps the style short definition code to the internal ID
		Map<String, Integer> numberFormatMap = new HashMap<String, Integer>(); // maps number format string to created POI format code
		if (workbook != null)
			ret.mapCellAddressToStyleId = new HashMap<Pair<String, CellAddress>, Integer>();
		int nextFreeStyleId = 0;
		
		for (Map.Entry<String, SheetState> sheetStateEntry : xlsf.sheetStates.entrySet())
			for (CellAddress cellAddress : sheetStateEntry.getValue().cells.keySet()) {
				try {
					XlsFormatterState.CellState state = sheetStateEntry.getValue().cells.get(cellAddress);
					
					// check whether the cell state requires any formatting:
					String cellShortString = state.cellFormatToShortString(true, true); 
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
					
					int currentStyleId = -1;
					if (styleCodeToInternalIdMap.containsKey(cellShortString)) // this style has been seen before and an internal ID hence exists
						currentStyleId = styleCodeToInternalIdMap.get(cellShortString);
					else { // this style needs to be newly created
						currentStyleId = nextFreeStyleId++;
						styleCodeToInternalIdMap.put(cellShortString, currentStyleId);
						styleIdToCreationInstructionMap.put(currentStyleId, new StyleCreationInstruction(state, font, numberFormatCode));
					}
					if (workbook != null)
						ret.mapCellAddressToStyleId.put(Pair.of(sheetStateEntry.getKey(), cellAddress), currentStyleId);
				}
				catch (Exception e) {
					StringWriter sw = new StringWriter();
					PrintWriter pw = new PrintWriter(sw);
					e.printStackTrace(pw);
					if (logger != null)
						logger.error(e.getClass().getCanonicalName() + ": " + e.getMessage() + "\n" + sw.toString());
					throw new Exception(cellAddress.formatAsString() + ": " + e.getClass().getCanonicalName() + ":" + e.getMessage(), e);
				}
				if (exec != null)
					exec.checkCanceled();
			}
		
		// fill the XLS artifact counts via the argument side channel (+ 1 each since the unformatted XLS file already had 1 default style etc.)
		ret.xlsArtifactCount = new XlsArtifactCount();
		ret.xlsArtifactCount.StyleCount = styleIdToCreationInstructionMap.size() + 1;
		ret.xlsArtifactCount.FontCount = fontMap.size() + 1;
		ret.xlsArtifactCount.NumberFormatCount = numberFormatMap.size() + 1;
		
		// actually create the styles:
		if (workbook != null) {
			ret.mapStyleIdToPoiStyle = new HashMap<Integer, XSSFCellStyle>();
			try {
				int i = 1;
				for (Integer styleInternalId : styleIdToCreationInstructionMap.keySet()) {
					ProgressTools.showProgressText(exec, "Creating XLS style", i++, ret.xlsArtifactCount.StyleCount, null);
					StyleCreationInstruction instruction = styleIdToCreationInstructionMap.get(styleInternalId);
					ret.mapStyleIdToPoiStyle.put(styleInternalId, XlsFormatterApplyXlsfToPoiConversions.createCellStyle(workbook, instruction.cellState, instruction.font, instruction.numberFormatCode));
					if (exec != null)
						exec.checkCanceled();
				}
			}
			catch (Exception e) {
				StringWriter sw = new StringWriter();
				PrintWriter pw = new PrintWriter(sw);
				e.printStackTrace(pw);
				if (logger != null)
					logger.error(e.getClass().getCanonicalName() + ": " + e.getMessage() + "\n" + sw.toString());
				throw new Exception("Could not create POI style. " + e.getClass().getCanonicalName() + ":" + e.getMessage(), e);
			}
		}
				
		return ret;
	}
	
	/**
	 * Simulates an application of a XlsFormatterState and calculates the number of required XLS artifacts (such as styles).
	 * Returns the essential information as a concise message.
	 */
	public static String getDerivedStyleComplexityMessage(XlsFormatterState state, final ExecutionContext exec, final NodeLogger logger) {
		
		StyleAnalysisResult res = null;
		
		try {
			res = deriveNecessaryStyles(null, state, exec, logger); // calling with workbook==null means that only xlsArtifactCount will be populated in the returned analysisResult
		} catch (Exception e) {
			return "Error during calculation of required instructions / styles.";
		}
		
		String ret = "XLS Formatter port object load state: " + res.xlsArtifactCount.StyleCount + " additional of " + XlsFormattingStateValidator.MAX_CELL_STYLES_PER_WORKBOOK + " (" + Math.round(res.xlsArtifactCount.StyleCount / (double)XlsFormattingStateValidator.MAX_CELL_STYLES_PER_WORKBOOK * 100) + "%) allowed total cell styles";
		if (res.xlsArtifactCount.FontCount > 1)
			ret += ", " + res.xlsArtifactCount.FontCount + " additional of " + XlsFormattingStateValidator.MAX_FONTS_PER_WORKBOOK + " allowed total fonts";
		if (res.xlsArtifactCount.NumberFormatCount > 1)
			ret += ", " + res.xlsArtifactCount.NumberFormatCount + " additional of " + XlsFormattingStateValidator.MAX_NUMBER_FORMATS_PER_WORKBOOK + " allowed total number formats";

		return ret;
	}
	
	/**
	 * Simulates an application of a XlsFormatterState and calculates the number of required XLS artifacts (such as styles).
	 * Throws exceptions or sets warnings if too many artifacts are needed.
	 */
	public static void checkDerivedStyleComplexity(XlsFormatterState state, WarningMessageContainer warningMessageContainer,
			final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		StyleAnalysisResult res = null;
		res = deriveNecessaryStyles(null, state, exec, logger); // calling with workbook==null means that only xlsArtifactCount will be populated in the returned analysisResult
		
		if (res.xlsArtifactCount.StyleCount > XlsFormattingStateValidator.MAX_CELL_STYLES_PER_WORKBOOK - ALLOWED_PREVIOUS_STYLES)
			throw new Exception("The XLS Formatter port object has been loaded with more instructions / cell styles (" + res.xlsArtifactCount.StyleCount + ") than can be implemented in an XLS workbook (" + XlsFormattingStateValidator.MAX_CELL_STYLES_PER_WORKBOOK + ").");
		
		if (res.xlsArtifactCount.FontCount > XlsFormattingStateValidator.MAX_FONTS_PER_WORKBOOK)
			throw new Exception("The XLS Formatter port object has been loaded with more required fonts (" + res.xlsArtifactCount.FontCount + ") than can be implemented in an XLS workbook (" + XlsFormattingStateValidator.MAX_FONTS_PER_WORKBOOK + ").");
		
		if (res.xlsArtifactCount.NumberFormatCount > XlsFormattingStateValidator.MAX_NUMBER_FORMATS_PER_WORKBOOK)
			throw new Exception("The XLS Formatter port object has been loaded with more required number formats (" + res.xlsArtifactCount.NumberFormatCount + ") than can be implemented in an XLS workbook (" + XlsFormattingStateValidator.MAX_NUMBER_FORMATS_PER_WORKBOOK + ").");
		
		
		double requiredStyleQuotaUsage = res.xlsArtifactCount.StyleCount / (double)XlsFormattingStateValidator.MAX_CELL_STYLES_PER_WORKBOOK;
		if (requiredStyleQuotaUsage >	WARNING_THRESHOLD_QUOTA && warningMessageContainer != null)
			warningMessageContainer.addMessage("The XLS Formatter port object has already been loaded with " + Math.round(requiredStyleQuotaUsage * 100) + "% of the allowed maximum number of cell styles.");
		
		logger.debug("XLS Formatter port object validation: " + res.xlsArtifactCount.StyleCount + " of " + XlsFormattingStateValidator.MAX_CELL_STYLES_PER_WORKBOOK + " (" + Math.round(requiredStyleQuotaUsage * 100) + "%) allowed cell styles. Fonts: " + res.xlsArtifactCount.FontCount + " of " + XlsFormattingStateValidator.MAX_FONTS_PER_WORKBOOK + ". Number formats: " + res.xlsArtifactCount.NumberFormatCount + " of " + XlsFormattingStateValidator.MAX_NUMBER_FORMATS_PER_WORKBOOK + ".");
	}
	
	/**
	 * Converts the column width shown in a standard spreadsheet to that required by POI.
	 * Depends on some font setting, we assume standard here (written by POI, read by POI).
	 */
	private static int xlsToPoiStandardColumnWidthConversion(double xlsColumnWidth) {
		return (int)Math.round(257.5 * xlsColumnWidth + 165.37); // formula based on own experiment with different settings and linear regression on that
	}
	
	public static double poiToXlsColumnWidthConversion(int poiColumnWidth) {
		return (poiColumnWidth - 165.37d) / 257.5d;
	}
}
