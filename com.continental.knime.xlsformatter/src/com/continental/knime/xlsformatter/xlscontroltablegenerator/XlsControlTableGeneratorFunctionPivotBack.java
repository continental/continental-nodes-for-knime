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

package com.continental.knime.xlsformatter.xlscontroltablegenerator;

import java.util.HashMap;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.lang3.mutable.MutableObject;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.MissingCell;
import org.knime.core.data.RowKey;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.IntCell;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.BufferedDataContainer;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeLogger;

import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableCreateTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;
import com.continental.knime.xlsformatter.commons.XlsFormatterTagTools;
import com.continental.knime.xlsformatter.xlscontroltablegenerator.XlsControlTableGeneratorNodeModel.InconsistencyResolutionOptions;

public class XlsControlTableGeneratorFunctionPivotBack {

	/* Note that this function cannot provide a meaningful spec without having the full table data yet, as only
	 * then it is known how many columns the resulting XLS Control Table will have. */
	
	private static Pattern regexPatternCell = Pattern.compile("^[A-Z]{1,3}[0-9]{1,7}$");
	private static Pattern regexPatternColumn = Pattern.compile("^[A-Z]{1,3}$");
	private static Pattern regexPatternColumnComparable = Pattern.compile("^(00[A-Z]|0[A-Z]{2}|[A-Z]{3})$");
	
	/**
	 * Transforms a long/unpivoted table into a wide XLS Control Table.
	 * Precondition: correct (i.e. pre-checked) column specification (incl. types)
	 */
	public static BufferedDataTable[] pivotBack(final BufferedDataTable dataTable,
			WarningMessageContainer warningMessageContainer,
			final XlsControlTableGeneratorNodeModel.InconsistencyResolutionOptions inconsistencyStrategy,
			final ExecutionContext exec, final NodeLogger logger) throws Exception {

		/* note that all cell address indices (row & col) are 0-based */
		
		
		// prepare data structures:
		Map<CellAddress, String> cellValues = new HashMap<CellAddress, String>(); 
		int maxTargetRow = -1;
		int maxTargetCol = -1;
		boolean showWarningRowInconsistencyPreferenceIsMissing = false;
		boolean showWarningColumnInconsistencyPreferenceIsMissing = false;
		
		
		// process input table:
		String[] originalColumnNames = dataTable.getSpec().getColumnNames();
		boolean hasSeenInvalidTagList = false;
		for (DataRow row : dataTable) {
			exec.checkCanceled();
			int targetRow = -1; // if the -1 is never overwritten (e.g. due to all missing values), an exception needs to be thrown
			int targetCol = -1; // if the -1 is never overwritten (e.g. due to all missing values), an exception needs to be thrown
			MutableObject<HiddenInconsistencyDetectionState> hiddenInconsistencyDetectionState = new MutableObject<HiddenInconsistencyDetectionState>();
			hiddenInconsistencyDetectionState.setValue(HiddenInconsistencyDetectionState.NOTHING_SEEN_YET);
			RowWarningState rowWarningState = RowWarningState.NOTHING_SEEN_YET;
			String targetValue = null;
			for (int c = 0; c < originalColumnNames.length; c++) {
				if (!row.getCell(c).isMissing())
					switch (originalColumnNames[c]) {
					case "Cell":
						if (!regexPatternCell.matcher(row.getCell(c).toString()).matches())
							throw new Exception("\"Cell\" value of \"" + row.getCell(c).toString() + "\" is not valid, must be like e.g. A1 or AB23. RowID " + row.getKey().getString());
						CellAddress addr = new CellAddress(row.getCell(c).toString());
						targetRow = setRowWithInconsistencyCheck(addr.getRow(), targetRow, row.getKey().getString(), originalColumnNames[c], inconsistencyStrategy);
						targetCol = setColumnWithInconsistencyCheck(addr.getColumn(), targetCol, row.getKey().getString(), originalColumnNames[c], hiddenInconsistencyDetectionState, inconsistencyStrategy);
						rowWarningState = rowWarningState == RowWarningState.SAW_ROW ? RowWarningState.SAW_ROW_AND_CELL : RowWarningState.SAW_CELL;
						break;
					case "Column":
						if (!regexPatternColumn.matcher(row.getCell(c).toString()).matches())
							throw new Exception("\"Column\" value of \"" + row.getCell(c).toString() + "\" is not valid, must be like e.g. A or AB. RowID " + row.getKey().getString());
						targetCol = setColumnWithInconsistencyCheck(
								CellReference.convertColStringToIndex(row.getCell(c).toString()),
								targetCol, row.getKey().getString(), originalColumnNames[c], hiddenInconsistencyDetectionState, inconsistencyStrategy);
						break;
					case "Column (comparable)":
						String colComparable = row.getCell(c).toString();
						if (colComparable.length() != 3)
							throw new Exception("\"Column (comparable)\" must have a length of 3 (e.g. 00A or ABC). RowID " + row.getKey().getString());
						if (!regexPatternColumnComparable.matcher(colComparable).matches())
							throw new Exception("\"Column (comparable)\" value of \"" + colComparable + "\" is not valid, must be like e.g. 00A or ABC. RowID " + row.getKey().getString());
						targetCol = setColumnWithInconsistencyCheck(
								CellReference.convertColStringToIndex(colComparable.replace("0", "")),
								targetCol, row.getKey().getString(), originalColumnNames[c], hiddenInconsistencyDetectionState, inconsistencyStrategy);
						break;
					case "Column (number)":
						int colNumber = ((IntCell)row.getCell(c)).getIntValue();
						if (colNumber <= 0)
							throw new Exception("Negative or zero \"Column (number)\" is not allowed (RowID " + row.getKey().getString() + ")");
						targetCol = setColumnWithInconsistencyCheck(
								colNumber - 1, targetCol, row.getKey().getString(), originalColumnNames[c], hiddenInconsistencyDetectionState, inconsistencyStrategy);
						break;
					case "Row":
						int rowNumber = ((IntCell)row.getCell(c)).getIntValue();
						if (rowNumber <= 0)
							throw new Exception("Negative or zero \"Row\" is not allowed (RowID " + row.getKey().getString() + ")");
						targetRow = setRowWithInconsistencyCheck(
								rowNumber - 1, targetRow, row.getKey().getString(), originalColumnNames[c], inconsistencyStrategy);
						rowWarningState = rowWarningState == RowWarningState.SAW_CELL ? RowWarningState.SAW_ROW_AND_CELL : RowWarningState.SAW_ROW;
						break;
					case "Value":
						targetValue = row.getCell(c).toString();
						hasSeenInvalidTagList |= !XlsFormatterTagTools.isValidTagList(targetValue);
						break;
					default:
					}
			}
			
			// handle missing target dimension in entire input table row:
			if (targetCol == -1)
				throw new Exception("Completely missing target column information in RowID " + row.getKey().getString());
			if (targetRow == -1)
				throw new Exception("Completely missing target row information in RowID " + row.getKey().getString());
			
			// handle errors and warnings regarding inconsistencies:
			if (hiddenInconsistencyDetectionState.getValue() == HiddenInconsistencyDetectionState.INCONSISTENCY_SEEN_BUT_NOT_YET_SUPERIOR_COLUMN)
				throw new Exception("The chosen contradiction resolution column is missing, but the other columns hold contradicting target column information. RowID " + row.getKey().getString());
			
			if (inconsistencyStrategy != InconsistencyResolutionOptions.FAIL) { // because in the other case, there is no superior column and hence no corresponding warning that this was missing
				showWarningRowInconsistencyPreferenceIsMissing |=
						(inconsistencyStrategy == InconsistencyResolutionOptions.CELL && rowWarningState == RowWarningState.SAW_ROW) ||
						(inconsistencyStrategy != InconsistencyResolutionOptions.CELL && rowWarningState == RowWarningState.SAW_CELL);
				
				showWarningColumnInconsistencyPreferenceIsMissing |= hiddenInconsistencyDetectionState.getValue() ==
						HiddenInconsistencyDetectionState.VALUE_FROM_INFERIOR_COLUMN; // superior column was never seen, but also no inconsistency
			}
			
			// update max sheet dimension trackers:
			if (targetCol > maxTargetCol)
				maxTargetCol = targetCol;
			if (targetRow > maxTargetRow)
				maxTargetRow = targetRow;
			
			// store target value in cell value map: 
			CellAddress targetAddress = new CellAddress(targetRow, targetCol);
			if (cellValues.containsKey(targetAddress))
				throw new Exception("Two input table rows target the same output table cell " + targetAddress.formatAsString() + ". This is not supported. Please GroupBy your input table first.");
			cellValues.put(targetAddress, targetValue);
		}
		
		// show warnings:
		if (!XlsFormatterControlTableValidator.isSheetSizeWithinXlsSpec(
				maxTargetCol + 1, maxTargetRow + 1))
			throw new Exception("The resulting XLS Control Table would exceed the maximum allowed size of a XLS sheet.");
		if (showWarningRowInconsistencyPreferenceIsMissing)
			warningMessageContainer.addMessage("The chosen row contradiction resolution preference could not always be followed due to missing cell(s).");
		if (showWarningColumnInconsistencyPreferenceIsMissing)
			warningMessageContainer.addMessage("The chosen column contradiction resolution preference could not always be followed due to missing cell(s), but the remaining columns held consistent information.");
		if (hasSeenInvalidTagList) {
			warningMessageContainer.addMessage("The generated table is not a fully valid XLS Formatter Control Table as it contains invalid characters in tags. See log for details.");
			logger.warn("Cells of a XLS Formatter Control Table shall contain comma-separated lists of tags. Tags are typically user chosen and speaking names, e.g. 'header' or 'totals'. Valid tags do not contain any of the letters '" + XlsFormatterTagTools.INVALID_TAGLIST_CHARACTERS + "'. This warning can be ignored for some special features, e.g. the 'applies to all tags' option of the XLS Border Formatter node.");
		}
		
		// generate output structure:
		DataTableSpec dataTableSpec = XlsFormatterControlTableCreateTools.createDataTableSpec(maxTargetCol + 1);
		BufferedDataContainer outBuffer = exec.createDataContainer(dataTableSpec);
		for (int r = 0; r <= maxTargetRow; r++) {
			DataCell[] cells = new DataCell[maxTargetCol + 1];
			for (int c = 0; c <= maxTargetCol; c++) {
				CellAddress addr = new CellAddress(r, c);
				if (cellValues.containsKey(addr)) {
					String value = cellValues.get(addr);					
					cells[c] = value == null ? new MissingCell(null) : new StringCell(value);
				}
				else
					cells[c] = new MissingCell(null);
			}
			outBuffer.addRowToTable(new DefaultRow(new RowKey(Integer.toString(r + 1)), cells));
		}
		outBuffer.close();
		return new BufferedDataTable[] { outBuffer.getTable() };
	}

	
	/**
	 * State of a state machine to detect whether a warning needs to be shown regarding missing superior row
	 */
	private enum RowWarningState {
		NOTHING_SEEN_YET,
		SAW_CELL,
		SAW_ROW,
		SAW_ROW_AND_CELL
	}
	
	/**
	 * Checks row assignment in light of the chosen inconsistency resolution strategy
	 * @param newRow The new value that shall be set.
	 * @param oldRow The value that is about to be overwritten.
	 * @param rowId The rowID as identifier of the problem source for a potential exception message.
	 * @param columnName The original table's column name, e.g. Cell, Column (number)
	 * @param option The inconsistency resolution strategy to be applied.
	 * @return The value to be set.
	 */
	private static int setRowWithInconsistencyCheck(int newRow, int oldRow, String rowId,
			String columnName, InconsistencyResolutionOptions option) throws Exception {
		if (oldRow == -1 || newRow == oldRow)
			return newRow;
		if (option == InconsistencyResolutionOptions.FAIL)
			throw new Exception("Contradicting target row information in RowID \"" + rowId + "\": " + (oldRow + 1) + " vs. " + (newRow + 1) + ".");
		if ((columnName.equals("Cell") && option == InconsistencyResolutionOptions.CELL) ||
				(columnName.equals("Row") && option != InconsistencyResolutionOptions.CELL))
			return newRow;
		return oldRow;
	}
	
	/**
	 * State of a state machine to detect inconsistencies despite the superior column being missing
	 */
	private enum HiddenInconsistencyDetectionState {
		NOTHING_SEEN_YET,
		VALUE_FROM_INFERIOR_COLUMN,
		SUPERIOR_COLUMN_SEEN,
		INCONSISTENCY_SEEN_BUT_NOT_YET_SUPERIOR_COLUMN
	}
	
	/**
	 * Checks column assignment in light of the chosen inconsistency resolution strategy
	 * @param newCol The new value that shall be set. Cannot be -1, only valid, non-missing column ids are allowed.
	 * @param oldCol The value that is about to be overwritten. Can be -1 for 'not yet defined'.
	 * @param rowId The rowID as identifier of the problem source for a potential exception message.
	 * @param columnName The original table's column name, e.g. Cell, Column (number)
	 * @param detectionState The mutable HiddenInconsistencyDetectionState 
	 * @param option The inconsistency resolution strategy to be applied.
	 * @return The value to be set.
	 */
	private static int setColumnWithInconsistencyCheck(int newCol, int oldCol, String rowId,
			String columnName, MutableObject<HiddenInconsistencyDetectionState> detectionState, InconsistencyResolutionOptions option) throws Exception {
		
		boolean isSuperiorColumn =
				(columnName.equals("Cell") && option == InconsistencyResolutionOptions.CELL) ||
				(columnName.equals("Column") && option == InconsistencyResolutionOptions.COLUMN) ||
				(columnName.equals("Column (comparable)") && option == InconsistencyResolutionOptions.COLUMN_COMPARABLE) ||
				(columnName.equals("Column (number)") && option == InconsistencyResolutionOptions.COLUMN_NUMBER);
		
		if (oldCol == -1) { // this is the first non-missing cell in this row
			detectionState.setValue(isSuperiorColumn ?
					HiddenInconsistencyDetectionState.SUPERIOR_COLUMN_SEEN :
					HiddenInconsistencyDetectionState.VALUE_FROM_INFERIOR_COLUMN);
			return newCol;
		}
		
		if (isSuperiorColumn)
			detectionState.setValue(HiddenInconsistencyDetectionState.SUPERIOR_COLUMN_SEEN);
		
		if (newCol == oldCol || isSuperiorColumn)
			return newCol;
		
		if (option == InconsistencyResolutionOptions.FAIL) // reaching this point means there is an inconsistency
			throw new Exception("Contradicting target column information in RowID \"" + rowId + "\": " + (oldCol + 1) + " (" + CellReference.convertNumToColString(oldCol) + ") vs. " + (newCol + 1) + " (" + CellReference.convertNumToColString(newCol) + ") .");
		
		if (detectionState.getValue() != HiddenInconsistencyDetectionState.SUPERIOR_COLUMN_SEEN)
			detectionState.setValue(HiddenInconsistencyDetectionState.INCONSISTENCY_SEEN_BUT_NOT_YET_SUPERIOR_COLUMN);
		
		return oldCol; // because a previous column might have been superior already
	}
}
