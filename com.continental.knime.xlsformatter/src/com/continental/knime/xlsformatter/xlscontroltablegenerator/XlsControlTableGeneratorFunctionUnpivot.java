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

import org.apache.poi.ss.util.CellReference;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataType;
import org.knime.core.data.MissingCell;
import org.knime.core.data.RowKey;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.IntCell;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.BufferedDataContainer;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeLogger;

import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;

public class XlsControlTableGeneratorFunctionUnpivot {

	private final static String[] columnNames = new String[] { "Column", "Row", "Value" };
	private final static String[] columnNamesExtended = new String[] { "Cell", "Column", "Column (comparable)", "Column (number)", "Column name", "Row", "RowID", "Value" };
	private final static DataType[] columnDataTypes = new DataType[] { StringCell.TYPE, IntCell.TYPE, StringCell.TYPE };
	private final static DataType[] columnDataTypesExtended = new DataType[] {StringCell.TYPE, StringCell.TYPE, StringCell.TYPE, IntCell.TYPE, StringCell.TYPE, IntCell.TYPE, StringCell.TYPE, StringCell.TYPE };

	public static DataTableSpec getUnpivotSpec(final DataTableSpec incomingSpec, final boolean addExtraHeaderColumns) {
		return new DataTableSpec(DataTableSpec.createColumnSpecs(
				addExtraHeaderColumns ? columnNamesExtended : columnNames,
						addExtraHeaderColumns ? columnDataTypesExtended : columnDataTypes));
	}

	public static BufferedDataTable[] unpivot(final BufferedDataTable dataTable, final boolean columnHeaderToFirstRow,
			final ExecutionContext exec, final NodeLogger logger, final boolean addExtraHeaderColumns) throws Exception {

		// prepare control table's column definition
		String[] originalColumnNames = dataTable.getSpec().getColumnNames();
		DataTableSpec dataTableSpec = getUnpivotSpec(dataTable.getSpec(), addExtraHeaderColumns);

		if (!XlsFormatterControlTableValidator.isSheetSizeWithinXlsSpec(
				originalColumnNames.length, dataTable.size() + (columnHeaderToFirstRow ? 1 : 0)))
			throw new Exception("Input table is bigger than the maximum allowed size of an XLS sheet.");

		// create output table
		BufferedDataContainer outBuffer = exec.createDataContainer(dataTableSpec);
		int currentRow = 0;
		long outputRowCounter = 0L; // note: much longer as an output row represents an input cell in this unpivot operation

		// add input table column header, if desired:
		if (columnHeaderToFirstRow) {
			for (int c = 0; c < originalColumnNames.length; c++) {
				exec.checkCanceled();
				DataCell[] cells = processInputCell(c, originalColumnNames[c], currentRow, null, originalColumnNames[c], addExtraHeaderColumns);
				outBuffer.addRowToTable(new DefaultRow(RowKey.createRowKey(outputRowCounter++), cells));
			}
			currentRow++;
		}

		// add input table cells:
		for (DataRow row : dataTable) {
			for (int c = 0; c < originalColumnNames.length; c++) {
				exec.checkCanceled();
				String value = row.getCell(c).isMissing() ? null : row.getCell(c).toString();
				DataCell[] cells = processInputCell(c, originalColumnNames[c], currentRow, row.getKey().toString(), value, addExtraHeaderColumns);
				outBuffer.addRowToTable(new DefaultRow(RowKey.createRowKey(outputRowCounter++), cells));
			}
			currentRow++;
		}
		outBuffer.close();
		return new BufferedDataTable[] { outBuffer.getTable() };
	}

	/**
	 * Transforms an input table cell to an (unpivoted) output table row (resp. its DataCell array).
	 * @param cellColIndex 0-based column index
	 * @param originalColumnName the name of the cell's input tabe column
	 * @param cellRowIndex 0-based row index
	 * @param rowId The KNIME RowID String of the cell's input table row
	 * @param value Null in case of a missing cell, the String value of the cell otherwise.
	 * @param addExtraHeaderColumns Option to include extra columns or just the basic XLS addressing columns.
	 * @return The data cell array that constitutes the output table row for this input table cell.
	 */
	private static DataCell[] processInputCell(final int cellColIndex, final String originalColumnName,
			final int cellRowIndex, final String rowId, final String value, final boolean addExtraHeaderColumns) {
		String colString = CellReference.convertNumToColString(cellColIndex);
		DataCell[] cells = new DataCell[addExtraHeaderColumns ? columnNamesExtended.length : columnNames.length];
		DataCell valueCell = value == null ? new MissingCell(null) : new StringCell(value);
		if (addExtraHeaderColumns) {
			cells[0] = new StringCell(colString + (cellRowIndex + 1));
			cells[1] = new StringCell(colString);
			cells[2] = new StringCell(XlsFormatterControlTableValidator.padColumnName(colString));
			cells[3] = new IntCell(cellColIndex + 1);
			cells[4] = new StringCell(originalColumnName);
			cells[5] = new IntCell(cellRowIndex + 1);
			cells[6] = rowId == null ? new MissingCell(null) : new StringCell(rowId);
			cells[7] = valueCell;
		}
		else {
			cells[0] = new StringCell(colString);
			cells[1] = new IntCell(cellRowIndex + 1);
			cells[2] = valueCell;
		};
		return cells;
	}
}
