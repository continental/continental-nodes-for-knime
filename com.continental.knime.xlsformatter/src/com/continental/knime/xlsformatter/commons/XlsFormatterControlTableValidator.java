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

package com.continental.knime.xlsformatter.commons;

import org.apache.poi.ss.util.CellReference;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataColumnSpec;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.StringValue;
import org.knime.core.data.def.DoubleCell;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeLogger;

public class XlsFormatterControlTableValidator {

	public static final int XLS_SIZE_LIMIT_MAX_ROWS = 1048576;
	public static final int XLS_SIZE_LIMIT_MAX_COLUMNS = 16384;
	
	/**
	 * The standard control table is STRING, meaning all column headers are of type String. For some nodes, slight modifications
	 * are allowed. E.g. for Row and Column Sizer, a Double typed column header is allowed in order to receive data.
	 * For e.g. Cell Merger and Border Formatter, there is an "all tags" option instead entering a single tag in the UI which
	 * is then matched against the control table. This option shall deactivate cell content check for invalid characters. It
	 * is also applicable to the Cell Background Colorizer node which takes direct String-typed color codes. 
	 */
	public enum ControlTableType { STRING, DOUBLE, STRING_WITHOUT_CONTENT_CHECK }
	
	/**
	 * Checks whether the provided data table is conform to the XlsFormatter Control Table layout.
	 * All columns need to be of type String, rowkeys need to be 1, 2, ..., and columns need to be A, B, C, ... or 00A, 00B, 00C, ...
	 */
	public static boolean isControlTable(final BufferedDataTable dataTable, final ControlTableType type, 
			final ExecutionContext exec, final NodeLogger logger) throws CanceledExecutionException {
		
		// check table size spec
		DataTableSpec spec = dataTable.getSpec();
		
		// check column spec
		String columnSpecValidationErrorMessage = getControlTableSpecValidationErrorMessage(spec, type);
		if (columnSpecValidationErrorMessage != null) {
			logger.warn(columnSpecValidationErrorMessage);
			return false;
		}
		
		if (!isSheetSizeWithinXlsSpec(1, dataTable.size()))  // column count 1 because it has been checked in the spec validation above already
		{
			logger.warn("Xls Formatting Control Table check failed: XLS maximum row limit violated.");
			return false;
		}
		
		// check row IDs and cell contents for invalid characters
		long currentRow = 0;
		for (DataRow row : dataTable) {
			exec.checkCanceled();
			int rowIndex;
			try {
				rowIndex = Integer.parseInt(row.getKey().getString());
			} catch (NumberFormatException ne) {
				logger.warn("Xls Formatting Control Table check failed: rowID \"" + row.getKey().getString() + "\" cannot be parsed to Integer.");
				return false;
			}
			if (rowIndex != currentRow + 1) {
				logger.warn("Xls Formatting Control Table check failed due to unexpected row order: rowID " + rowIndex + " does not match the expected " + (currentRow + 1));
				return false;
			}
			currentRow++;
			
			if (type == ControlTableType.STRING)
				for (int c = 0; c < spec.getNumColumns(); c++) {
					DataCell cell = row.getCell(c);
					if (cell != null && !cell.isMissing() && !XlsFormatterTagTools.isValidTagList(cell.toString())) {
						logger.warn("Xls Formatting Control Table check failed at rowID \"" + row.getKey().getString() + "\", column \"" + spec.getColumnSpec(c).getName() + "\": Comma-separated tag list contains invalid character(s), i.e. " + XlsFormatterTagTools.INVALID_TAGLIST_CHARACTERS);
						return false;
					}
				}
		}
		
		// all checks passed successfully, so return true
		return true;
	}
	
	/**
	 * Checks whether the provided data table is conform to the XlsFormatter Control Table layout.
	 * All columns need to be of type String, rowkeys need to be 1, 2, ..., and columns need to be A, B, C, ... or 00A, 00B, 00C, ...
	 */
	public static boolean isControlTable(final BufferedDataTable dataTable, 
			final ExecutionContext exec, final NodeLogger logger) throws CanceledExecutionException {
		return isControlTable(dataTable, ControlTableType.STRING, exec, logger);
	}
	
	/**
	 * Check whether a DataTableSpec represents a valid XLS Formatting control table header of
	 * a specified type (integer or string).
	 */
	public static boolean isControlTableSpec(final DataTableSpec spec, ControlTableType type, final NodeLogger logger) {
		String columnSpecValidationErrorMessage = getControlTableSpecValidationErrorMessage(spec, type);
		if (columnSpecValidationErrorMessage == null)
			return true;
		logger.warn(columnSpecValidationErrorMessage);
		return false;
	}
	
	/**
	 * Check whether a DataTableSpec represents a valid XLS Formatting control table header,
	 * determining the style based on the first column's data type (double or string) automatically.
	 */
	public static boolean isControlTableSpecDoubleOrString(final DataTableSpec spec, final NodeLogger logger) {
		
		if (spec.getNumColumns() == 0)
			return false;
  	ControlTableType type = XlsFormatterControlTableAnalysisTools.isDoubleControlTableSpecCandidate(spec) ? ControlTableType.DOUBLE : ControlTableType.STRING;
		
		String columnSpecValidationErrorMessage = getControlTableSpecValidationErrorMessage(spec, type);
		if (columnSpecValidationErrorMessage == null)
			return true;
		logger.warn(columnSpecValidationErrorMessage);
		return false;
	}
	
	/**
	 * Check whether a DataTableSpec represents a valid XLS Formatting control table header.
	 */
	public static boolean isControlTableSpec(final DataTableSpec spec, final NodeLogger logger) {
		return isControlTableSpec(spec, ControlTableType.STRING, logger);
	}
	
	/**
	 * Checks whether a data table spec is a valid XLS Formatting control table header.
	 * @return An error message of why the spec is invalid or null if it is valid.
	 */
	private static String getControlTableSpecValidationErrorMessage(final DataTableSpec spec, final ControlTableType type) {
		
		// check table size spec
		int colCount = spec.getNumColumns();
		if (!isSheetSizeWithinXlsSpec(colCount, 1)) // row 1 because row is uncrucial for this column header check
			return "Xls Formatting Control Table check failed: XLS maximum column limit violated.";
		
		// check column header
		int initialColumnNameLength = 0; // used to detect consistency in padding usage
		for (int c = 0; c < colCount; c++) {
  		DataColumnSpec columnSpec = spec.getColumnSpec(c);
  		if ((type == ControlTableType.STRING || type == ControlTableType.STRING_WITHOUT_CONTENT_CHECK) && !columnSpec.getType().isCompatible(StringValue.class))
  			return "Xls Formatting Control Table check failed: all columns must be of type String. Possible solution: Use Number to String node.";
  		if (type == ControlTableType.DOUBLE && !columnSpec.getType().getCellClass().equals(DoubleCell.class))
  			return "Xls Formatting Control Table check failed: all columns must be of type Double for usage with this node.";
  		String expectedXlsColName = CellReference.convertNumToColString(c);
  		String expectedPaddedXlsColName = XlsFormatterControlTableValidator.padColumnName(expectedXlsColName);
  		if (!expectedXlsColName.equals(columnSpec.getName())
  				&& !expectedPaddedXlsColName.equals(columnSpec.getName()))
  			return "Xls Formatting Control Table check failed: Column name " + columnSpec.getName() + " is not matching the expected " + expectedXlsColName + " or " + expectedPaddedXlsColName + ".";
  		if (c == 0)
  			initialColumnNameLength = columnSpec.getName().length();
  		else if ((initialColumnNameLength == 3 && !columnSpec.getName().equals(expectedPaddedXlsColName))
  				|| (initialColumnNameLength == 1 && !columnSpec.getName().equals(expectedXlsColName)))
  			return "Xls Formatting Control Table check failed: Inconsistent usage of padding in column names (either all or no column names may be padded with leading zeros).";
  	}
		return null;
	}
	
	
	/**
	 * Pad a XLS column name with leading 0 (zero) characters to a total length of 3 characters.
	 * This makes two column names sortable and comparable in the expected way.
	 */
	public static String padColumnName(String colName) {
		String paddedCol = "00" + colName;
		return paddedCol.substring(paddedCol.length() - 3);
	}
	
	/**
	 * Checks whether a provided input table size is within the XLS specification regarding
	 * maximum column count and maximum row count.
	 */
	public static boolean isSheetSizeWithinXlsSpec(long colCount, long rowCount) {
		return rowCount <= XLS_SIZE_LIMIT_MAX_ROWS && colCount <= XLS_SIZE_LIMIT_MAX_COLUMNS;
	}
	
	/**
	 * Check whether a 0-based cell address is valid (i.e. on the sheet).
	 * Detects negative values and max size limit violations.
	 */
	public static boolean isCellWithinXlsSpec(int row, int column) {
		return 0 <= row && row < XLS_SIZE_LIMIT_MAX_ROWS && 0 <= column && column < XLS_SIZE_LIMIT_MAX_COLUMNS;
	}
}
