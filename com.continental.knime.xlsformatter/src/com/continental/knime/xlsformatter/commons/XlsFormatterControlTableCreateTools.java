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

import java.util.HashSet;
import java.util.Set;

import org.apache.poi.ss.util.CellReference;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataType;
import org.knime.core.data.MissingCell;
import org.knime.core.data.RowKey;
import org.knime.core.data.container.CloseableRowIterator;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.BufferedDataContainer;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeLogger;

public class XlsFormatterControlTableCreateTools {

	/**
	 * Generates a column specification conform to the Xls Formatting Control Table specification and derives a BufferedDataContainer from it.
	 */
	public static BufferedDataContainer getNewBufferedDataContainer(final int columnCount, final DataType dataType, final ExecutionContext exec, final NodeLogger logger) {	
		return exec.createDataContainer(createDataTableSpec(columnCount, dataType));
	}
	public static BufferedDataContainer getNewBufferedDataContainer(final int columnCount, final ExecutionContext exec, final NodeLogger logger) {	
		return exec.createDataContainer(createDataTableSpec(columnCount, StringCell.TYPE));
	}
	
	/**
	 * Creates a DataTableSpec of a XLS Formatting Control Table header of defined width.
	 */
	public static DataTableSpec createDataTableSpec(final int columnCount, final DataType dataType) {
		String[] newColumnNames = new String[columnCount];
		DataType[] newDatatypes = new DataType[columnCount];
		for (int c = 0; c < columnCount; c++) {
			newColumnNames[c] = CellReference.convertNumToColString(c);
			newDatatypes[c] = dataType;
		}
		return new DataTableSpec(DataTableSpec.createColumnSpecs(newColumnNames, newDatatypes));
	}
	public static DataTableSpec createDataTableSpec(final int columnCount) {
		return createDataTableSpec(columnCount, StringCell.TYPE);
	}
	
	/**
	 * Generates a RowKey array conform to the Xls Formatting Control Table specification.
	 */
	public static RowKey[] getRowkeyArray(final int rowCount, final ExecutionContext exec, final NodeLogger logger) {
		String[] rowKeys = new String[rowCount]; 
  	for (int r = 0; r < rowCount; r++)
  		rowKeys[r] = String.valueOf(r + 1);
  	return RowKey.toRowKeys(rowKeys);
	}
	
	/**
	 * Merge two XLS Formatter Control Tables into one. 
	 */
	public static BufferedDataTable merge(
			final BufferedDataTable primaryInputTable,
			final BufferedDataTable secondaryInputTable,
			final boolean overwriteInsteadOfAppend,
			final ExecutionContext exec, final NodeLogger logger) throws CanceledExecutionException {
		
		int topWidth = primaryInputTable.getSpec().getNumColumns();
  	int topHeight = (int)primaryInputTable.size();
  	int bottomWidth = secondaryInputTable.getSpec().getNumColumns();
  	int bottomHeight = (int)secondaryInputTable.size();
  	
  	int width = Math.max(topWidth, bottomWidth);
  	int height = Math.max(topHeight, bottomHeight);
  	
  	// prepare output table:
  	BufferedDataContainer outBuffer = XlsFormatterControlTableCreateTools.getNewBufferedDataContainer(width, exec, logger);
  	RowKey[] rowKeyArr = XlsFormatterControlTableCreateTools.getRowkeyArray(height, exec, logger);
  	
  	// fill output table:
  	try (CloseableRowIterator iteratorTop = primaryInputTable.iterator();
  			CloseableRowIterator iteratorBottom = secondaryInputTable.iterator()) {
    	for (int r = 0; r < height; r++) {
    		exec.checkCanceled();
    		DataRow rowTop;
    		DataRow rowBottom;
    		if (iteratorTop.hasNext())
    			rowTop = iteratorTop.next();
    		else
    			rowTop = null;
    		if (iteratorBottom.hasNext())
    			rowBottom = iteratorBottom.next();
    		else
    			rowBottom = null;
    		
  			DataCell[] cells = new DataCell[width];
  			for (int c = 0; c < width; c++) {
  				DataCell topCell = rowTop == null || topWidth <= c ? null : rowTop.getCell(c);
  				DataCell bottomCell = rowBottom == null || bottomWidth <= c ? null : rowBottom.getCell(c);
  				if (topCell != null && topCell.isMissing())
  					topCell = null;
  				if (bottomCell != null && bottomCell.isMissing())
  					bottomCell = null;
  				
  				boolean topCellHasContent = topCell != null && !topCell.toString().trim().equals("");
  				boolean bottomCellHasContent = bottomCell != null && !bottomCell.toString().trim().equals("");
  				
  				if ((!topCellHasContent && !bottomCellHasContent) ||
  						(overwriteInsteadOfAppend && bottomCell != null && bottomCell.toString().trim().equals(""))) // special deletion command
  					cells[c] = new MissingCell(null);
  				else {
  					String tags;
  					if (!topCellHasContent || (overwriteInsteadOfAppend && bottomCellHasContent))
  						tags = bottomCell.toString();
  					else if (!bottomCellHasContent)
  						tags = topCell.toString();
  					else
  						tags = topCell.toString() + "," + bottomCell.toString();
  					
  					cells[c] = new StringCell(removeDuplicateTags(tags)); // removeDuplicateTags performs a String.trim() call on each entry 
  				}
  			}
  			
  			DataRow rowOut = new DefaultRow(rowKeyArr[r], cells);
  			outBuffer.addRowToTable(rowOut);
    	}
  	}
  	outBuffer.close();
    return outBuffer.getTable();
	}
	
	/**
   * Removes duplicates from a comma-separated tag list and trims each element. Sort order is preserved.
   * @param value Comma-separated list of tags.
   * @return Comma-separated list of merged tags.
   */
  private static String removeDuplicateTags(String value) {
  	String[] tags = value.split(",");
  	Set<String> knownElements = new HashSet<String>();
  	StringBuilder sb = new StringBuilder();
  	for (String tag : tags) {
  		tag = tag.trim();
  		if (!knownElements.contains(tag)) {
  			if (knownElements.size() != 0)
  				sb.append(",");
  			sb.append(tag);
  			knownElements.add(tag);
  		}
  	}
  	return sb.toString();
  }
}