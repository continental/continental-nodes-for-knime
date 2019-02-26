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

import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.MissingCell;
import org.knime.core.data.RowKey;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.BufferedDataContainer;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeLogger;

import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellState;

public class AddressingTools {
	
	/**
	 * Parses an XLS range (can consist of a single cell as well), e.g. A1:D2, or R1C1:R2C2
	 */
	public static CellRangeAddress parseRange(String value) throws IllegalArgumentException {
		CellRangeAddress ret;
		Pattern r1c1RegexPattern = Pattern.compile("^R([0-9]{1,7})C([0-9]{1,5})(?:\\:R([0-9]{1,7})C([0-9]{1,5}))?$");
		Matcher match = r1c1RegexPattern.matcher(value);
		if (match.find()) { // regex matches, so we have a R1C1 syntax address which we need to parse ourself
			try {
				CellAddress cellFrom = validateR1C1Cell(match.group(1), match.group(2));
				if (match.groupCount() == 2 || match.group(3) == null)
					return new CellRangeAddress(cellFrom.getRow() - 1, cellFrom.getRow() - 1, cellFrom.getColumn() - 1, cellFrom.getColumn() - 1);
				CellAddress cellTo = validateR1C1Cell(match.group(3), match.group(4));
				return new CellRangeAddress(cellFrom.getRow() - 1, cellTo.getRow() - 1, cellFrom.getColumn() - 1, cellTo.getColumn() - 1);
			} catch (Exception e) {
				throw new IllegalArgumentException("Could not parse R1C1-syntaxed range \"" + value + "\". Please check whether your cell/row values are within the XLS sheet size limits and whether a range is specified as 'top-left cell, colon, bottom-right cell'.");
			}
			
		} else { // R1C1 syntax did not match via regex, so use POI to parse standard A1:B2 like syntax
			try {
				ret = CellRangeAddress.valueOf(value);
			}
			catch (Exception e) {
				throw new IllegalArgumentException("Could not parse range \"" + value + "\". Expected something like \"A1:D2\".");
			}
			
			if (ret.isFullColumnRange() || ret.isFullRowRange())
	  		throw new IllegalArgumentException("Full column or row ranges are not yet supported. Please specify a range from cell to cell.");
		}
		return ret;
	}
	
	/**
	 * Parse a single R1C1-syntax range coming from a regex match and check for conformance with sheet size limits.
	 * Returns null in case of failure.
	 */
	private static CellAddress validateR1C1Cell(String matchGroupRow, String matchGroupColumn) {
		try {
			int row = Integer.parseInt(matchGroupRow);
			int col = Integer.parseInt(matchGroupColumn);
			if (!AddressingTools.isCellAddressInSpec(col, row))
				return null;
			return new CellAddress(row, col);
		} catch (Exception e) {
			return null;
		}
	}
	
	/**
	 * Identifies (or creates) a cell in a map.
	 * @param cellMap The map of cell address and state that can but needs not contain the searched cell already.
	 * @param column The 0-based column id of the cell.
	 * @param row The 0-based row id of the cell.
	 * @return The searched cell's CellState or null, if the searched cell is outside the XLS sheet (e.g. left neighbor of A1).
	 */
	public static CellState safelyGetCellWithNegativeCheck(Map<CellAddress, CellState> cellMap, int column, int row) {
		if (!isCellAddressInSpec(column, row))
			return null;
		
		CellAddress searchedAddress = new CellAddress(row, column);
		return safelyGetCellInMap(cellMap, searchedAddress);
	}
	
	
	public static boolean isCellAddressInSpec(int column, int row) {
		return !((column < 0 || XlsFormatterControlTableValidator.XLS_SIZE_LIMIT_MAX_COLUMNS <= column ||
				row < 0 || XlsFormatterControlTableValidator.XLS_SIZE_LIMIT_MAX_ROWS <= row));
	}
	
	/**
	 * 
	 * @param cellsMap
	 * @param cell
	 * @return
	 */
	public static CellState safelyGetCellInMap(Map<CellAddress, CellState> cellsMap, CellAddress cell) {
		XlsFormatterState.CellState cellState;
		if (cellsMap.containsKey(cell))
			cellState = cellsMap.get(cell);
		else {
			cellState = new XlsFormatterState.CellState();
			cellsMap.put(cell, cellState);
		}
		return cellState;
	}
	
	/**
	 * Checks two lists of cell ranges for an overlap in any pair of ranges between both lists.
	 * Will NOT detect overlaps in pairs of one list.
	 */
	public static boolean hasOverlap(List<CellRangeAddress> list1, List<CellRangeAddress> list2, final ExecutionContext exec, final NodeLogger logger) throws CanceledExecutionException {
		for (CellRangeAddress range1 : list1)
			for (CellRangeAddress range2 : list2) {
				exec.checkCanceled();
				if (range1.intersects(range2)) {
					logger.warn("Cell range overlap detected between " + range1.formatAsString() + " and " + range2.formatAsString() + ".");
					return true;
				}
			}
		return false;
	}
	
	/**
	 * From a list of CellAddresses, determine rectangular ranges, which may be adjacent to each other. 
	 */
	public static List<CellRangeAddress> getRangesFromAddressList(final List<CellAddress> addresses, final ExecutionContext exec, final NodeLogger logger) throws Exception {
		String tag = "x";
		if (addresses.size() == 0)
			return new ArrayList<CellRangeAddress>(0);
		
		return XlsFormatterControlTableAnalysisTools.getRangesFromTag(
				getControlTableFromAddressList(addresses, tag, exec, logger), tag, false, true, null, exec, logger);
	}
	
	/**
	 * Gets the String length of the XLS formula that would represent a list of ranges with absolute positioning ($)
	 */
	public static int getFormulaLengthFromRangeList(List<CellRangeAddress> ranges) {
		int ret = 1; // for the = sign starting the formula
		for (CellRangeAddress range : ranges)
			ret += range.formatAsString().length() + (range.getNumberOfCells() == 1 ? 2 : 4) + 1; // inline if counts the number of dollar signs for absolute addressing, final 1 adder for closing semi-colon 
		return ret - 1; // -1 for removing the final semi-colon  
	}
	
	/**
	 * From a list of CellAddresses, create a matching control table.
	 */
	private static BufferedDataTable getControlTableFromAddressList(final List<CellAddress> addresses, final String tag, final ExecutionContext exec, final NodeLogger logger) throws Exception {
		addresses.sort(null); // use default row-major, column-minor sorting of CellAddress implementation
		exec.checkCanceled();
		int rowCount = addresses.get(addresses.size()-1).getRow() + 1;
		int columnCount = addresses.stream().map(ca -> ca.getColumn()).max(Comparator.comparing(Integer::valueOf)).get() + 1;
		exec.checkCanceled();
		Set<CellAddress> addressSet = new HashSet<CellAddress>(addresses);
		exec.checkCanceled();
		
		// prepare output table:
		BufferedDataContainer outBuffer = XlsFormatterControlTableCreateTools.getNewBufferedDataContainer(columnCount, exec, logger);
		RowKey[] rowKeyArr = XlsFormatterControlTableCreateTools.getRowkeyArray(rowCount, exec, logger);
		
		for (int r = 0; r < rowCount; r++) {
			DataCell[] cells = new DataCell[columnCount];
			for (int c = 0; c < columnCount; c++) {
				exec.checkCanceled();
				cells[c] = addressSet.contains(new CellAddress(r, c)) ? new StringCell(tag) : new MissingCell(null);
			}
			DataRow rowOut = new DefaultRow(rowKeyArr[r], cells);
			outBuffer.addRowToTable(rowOut);
		}
		outBuffer.close();
		return outBuffer.getTable();
	}
}
