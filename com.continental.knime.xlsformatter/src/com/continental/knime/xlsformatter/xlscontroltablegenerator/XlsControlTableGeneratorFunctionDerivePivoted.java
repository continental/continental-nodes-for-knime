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

import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.MissingCell;
import org.knime.core.data.RowKey;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.BufferedDataContainer;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeLogger;

import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableCreateTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;
import com.continental.knime.xlsformatter.commons.XlsFormatterTagTools;

public class XlsControlTableGeneratorFunctionDerivePivoted {

	public static DataTableSpec getPivotedSpec(final DataTableSpec incomingSpec) {
		return XlsFormatterControlTableCreateTools.createDataTableSpec(incomingSpec.getNumColumns());
	}
	
	public static BufferedDataTable[] derivePivoted(final BufferedDataTable dataTable, final boolean columnHeaderToFirstRow,
			final boolean copyContents, WarningMessageContainer warningMessageContainer,
			final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		// prepare control table's column definition
		int originalTableWidth = dataTable.getSpec().getNumColumns();
		
		if (!XlsFormatterControlTableValidator.isSheetSizeWithinXlsSpec(originalTableWidth, dataTable.size()))
  		throw new Exception("Input table is bigger than the maximum allowed size of an XLS sheet.");
		
  	BufferedDataContainer outBuffer = XlsFormatterControlTableCreateTools.getNewBufferedDataContainer(originalTableWidth, exec, logger);
  	RowKey[] rowKeyArr = XlsFormatterControlTableCreateTools.getRowkeyArray((int)dataTable.size() + (columnHeaderToFirstRow ? 1 : 0), exec, logger);
		
		int currentRow = 0;
		boolean hasSeenInvalidTagList = false; // in order to warn later on about an invalid XLS Formatter Control Table
		
		if (columnHeaderToFirstRow) {
			DataCell[] cells = new DataCell[originalTableWidth];
			String[] originalColumnNames = dataTable.getSpec().getColumnNames();
			for (int c = 0; c < originalTableWidth; c++) {
				cells[c] = !copyContents ? new MissingCell(null) : new StringCell(originalColumnNames[c]);
				hasSeenInvalidTagList |= copyContents && !XlsFormatterTagTools.isValidTagList(originalColumnNames[c]);
			}
			outBuffer.addRowToTable(new DefaultRow(rowKeyArr[currentRow++], cells));
		}
		
		for (DataRow row : dataTable) {
			exec.checkCanceled();
			DataCell[] cells = new DataCell[originalTableWidth];
			for (int c = 0; c < originalTableWidth; c++) {
				DataCell originalCell = row.getCell(c);
				cells[c] = !copyContents || originalCell.isMissing() ? new MissingCell(null) : new StringCell(originalCell.toString());
				hasSeenInvalidTagList |= copyContents && !originalCell.isMissing() && !XlsFormatterTagTools.isValidTagList(originalCell.toString());
			}			
			outBuffer.addRowToTable(new DefaultRow(rowKeyArr[currentRow++], cells));
		}
		outBuffer.close();
		
		if (hasSeenInvalidTagList) {
			warningMessageContainer.addMessage("Warning: The generated table is not a fully valid XLS Formatter Control Table as it contains invalid characters in tags. See log for details.");
			logger.warn("Cells of a XLS Formatter Control Table shall contain comma-separated lists of tags. Tags are typically user chosen and speaking names, e.g. 'header' or 'totals'. Valid tags do not contain any of the letters '" + XlsFormatterTagTools.INVALID_TAGLIST_CHARACTERS + "'. This warning can be ignored for some special features, e.g. the 'applies to all tags' option of the XLS Border Formatter node.");
		}
		
		return new BufferedDataTable[] { outBuffer.getTable() };
	}
}
