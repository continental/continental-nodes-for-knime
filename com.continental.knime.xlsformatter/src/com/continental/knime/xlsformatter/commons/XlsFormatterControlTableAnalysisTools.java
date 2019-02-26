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
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DoubleValue;
import org.knime.core.data.def.DoubleCell;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeLogger;

public class XlsFormatterControlTableAnalysisTools {

	/**
	 * Determines whether a spec could be that of a double-typed control table (based on the first column's data
	 * type only). Integer columns still qualify as candidate.
	 */
	public static boolean isDoubleControlTableSpecCandidate(DataTableSpec spec) {
		if (spec.getNumColumns() == 0)
			return false; 
		return spec.getColumnSpec(0).getType().isCompatible(DoubleValue.class);
	}
	
	/**
	 * Gets a list of cells matching a specified tag
	 */
	public static List<CellAddress> getCellsMatchingTag(
			final BufferedDataTable dataTable, final String tag, final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		List<CellAddress> ret = new ArrayList<CellAddress>();
		int colCount = dataTable.getSpec().getNumColumns();
		int r = 0;
		for (DataRow dataRow : dataTable) {
			exec.checkCanceled();
			for (int c = 0; c < colCount; c++) {
				DataCell cell = dataRow.getCell(c);
				String cellTags = cell == null || cell.isMissing() ? null : cell.toString();
				if (cellTags != null && XlsFormatterTagTools.doesTagMatch(cellTags, tag))
					ret.add(new CellAddress(r, c));
			}
			r++;
		}
		return ret;
	}
	
	/**
	 * Gets a list of CellAddress lists, especially useful for functionality that shall be executed on a
	 * pseudo-range for each tag appearing in the control table. Tag combination here means that the
	 * full String cell content is used, not the split result of separating the comma-separated tag list.
	 */
	public static List<List<CellAddress>> getCellsListForEachTagCombination(
			final BufferedDataTable dataTable, final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		List<List<CellAddress>> ret = new ArrayList<List<CellAddress>>();
		Map<String, List<CellAddress>> mapTagToList = new HashMap<String, List<CellAddress>>();
		int colCount = dataTable.getSpec().getNumColumns();
		int r = 0;
		for (DataRow dataRow : dataTable) {
			exec.checkCanceled();
			for (int c = 0; c < colCount; c++) {
				DataCell cell = dataRow.getCell(c);
				if (cell == null || cell.isMissing())
					continue;
				String tags = cell.toString();
				if (!mapTagToList.containsKey(tags)) {
					List<CellAddress> tempList = new ArrayList<CellAddress>();
					mapTagToList.put(tags, tempList);
					ret.add(tempList); // just add the reference to the list, it's contents are not final at this point in time yet
				}
				List<CellAddress> list = mapTagToList.get(tags);
				list.add(new CellAddress(r, c));
			}
			r++;
		}
		return ret;
	}
	
	/**
	 * Gets a list of all strings that appear in the table (will be used as tags, but without tag validation and splitting
	 * logic, pure cell content is used here).
	 */
	public static List<String> getAllFullCellContentTags(
			final BufferedDataTable dataTable, final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		Set<String> tags = new HashSet<String>();
		
		int colCount = dataTable.getSpec().getNumColumns();
		for (DataRow dataRow : dataTable) {
			exec.checkCanceled();
			for (int c = 0; c < colCount; c++) {
				DataCell cell = dataRow.getCell(c);
				if (cell == null || cell.isMissing())
					continue;
				tags.add(cell.toString().trim());
			}
		}
		return new ArrayList<String>(tags);
	}
	
	/**
	 * Gets a map of non-null cells and their String values
	 */
	public static Map<CellAddress, String> getCellsWithStringValues(
			final BufferedDataTable dataTable, final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		Map<CellAddress, String> ret = new HashMap<CellAddress, String>();
		int colCount = dataTable.getSpec().getNumColumns();
		int r = 0;
		for (DataRow dataRow : dataTable) {
			for (int c = 0; c < colCount; c++) {
				DataCell cell = dataRow.getCell(c);
				if (cell != null && !cell.isMissing())
					ret.put(new CellAddress(r, c), cell.toString());
			}
			r++;
		}
		return ret;
	}
	
	/**
	 * Gets a map of non-null cells and their double values
	 */
	public static Map<CellAddress, Double> getCellsWithDoubleValues(
			final BufferedDataTable dataTable, final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		Map<CellAddress, Double> ret = new HashMap<CellAddress, Double>();
		int colCount = dataTable.getSpec().getNumColumns();
		int r = 0;
		for (DataRow dataRow : dataTable) {
			for (int c = 0; c < colCount; c++) {
				DataCell cell = dataRow.getCell(c);
				if (cell != null && !cell.isMissing())
					ret.put(new CellAddress(r, c), ((DoubleCell)cell).getDoubleValue());
			}
			r++;
		}
		return ret;
	}
	
	/**
	 * Gets a list of rectangular cell ranges defined by a single tag.
	 * Adjacent ranges are not allowed to share any cell's border, just a cell corner.
	 * @param fullCellContentInsteadOfTag If true, the UI option of "all tags" is probably present and the full cell content shall be checked instead of finding a single tag in a comma-separated list of tags
	 * @param adjacentRangesAllowed If true, two ranges found for this tag may share cell borders, if false, they may only share cell corners.
	 */
	public static List<CellRangeAddress> getRangesFromTag (
			final BufferedDataTable dataTable, final String tag, boolean fullCellContentInsteadOfTag, boolean adjacentRangesAllowed,
			WarningMessageContainer warningMessage,	final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		int colCount = dataTable.getSpec().getNumColumns();
		
		// prepare variables for searching the entire sheet for rectangular areas:
		Map<Integer, CellRangeAddress> mapColumnToCurrentlyOpenRange = new HashMap<Integer, CellRangeAddress>();
		List<CellRangeAddress> ret = new ArrayList<CellRangeAddress>();

		int r = 0;
		boolean hasAnyMatch = false;
		for (DataRow dataRow : dataTable) {
			CellRangeAddress leftNeighborsNewRange = null; // only important for the top row of a range, subsequent rows are handled via the map
			for (int c = 0; c < colCount; c++) {
				DataCell dataCell = dataRow == null ? null : dataRow.getCell(c);
				String cellTags = dataCell == null || dataCell.isMissing() ? null : dataCell.toString();
				boolean isMatch = cellTags == null ? false :
					(fullCellContentInsteadOfTag ? cellTags.equals(tag) : XlsFormatterTagTools.doesTagMatch(cellTags, tag));
				hasAnyMatch |= isMatch;
				
				CellRangeAddress range = mapColumnToCurrentlyOpenRange.get(c);
				if (range == null && isMatch) {
					if (leftNeighborsNewRange == null) {
						range = new CellRangeAddress(r, r, c, c);
						leftNeighborsNewRange = range;
						ret.add(range);
					}
					else { // has a left neighbor already (but still not one above)
						range = leftNeighborsNewRange;
						range.setLastColumn(c);
					}
					mapColumnToCurrentlyOpenRange.put(c, range);
				}
				else if (range != null && isMatch) { // has one above
					if (range.getLastRow() != r && range.getFirstColumn() == c) { // if this is a new line in an existing block, extend the range
						range.setLastRow(r);
					}
					else if (range.getLastRow() != r) // we are at least one to the right of a new line start, but the left neighbor didn't extend the range's last row -> must be a non-rectangular range
						splitRange(mapColumnToCurrentlyOpenRange, ret, range, c, range.getLastRow(), r, adjacentRangesAllowed, tag);
					leftNeighborsNewRange = null; // because this column's range is extended from above, not growing from left to right on an range's first row
				}
				else if (range != null && !isMatch) { // finishes one above
					mapColumnToCurrentlyOpenRange.remove(c);
					if (r == range.getLastRow()) // a left neighbor has already extended this range to the current row
						splitRange(mapColumnToCurrentlyOpenRange, ret, range, c, r, r-1, adjacentRangesAllowed, tag);
					leftNeighborsNewRange = null; // because this column's range is extended from above, not growing from left to right on an range's first row
				}
				else // range == null && !isMatch  is the only option left
					leftNeighborsNewRange = null;
			}
			r++;
		}
		
		if (!hasAnyMatch && warningMessage != null)
			warningMessage.addMessage(TagBasedXlsCellFormatterNodeModel.getWarningMessage(tag));
		
		// check for adjacent ranges:
		List<CellRangeAddress> toDeleteOneCellRange = new ArrayList<CellRangeAddress>();
		if (!adjacentRangesAllowed) {
			for (CellRangeAddress range1 : ret) {
				for (CellRangeAddress range2 : ret)
					if (range1 != range2) {
						if ((range1.getFirstColumn() - 1 == range2.getLastColumn() && !(range1.getFirstRow() > range2.getLastRow() || range1.getLastRow() < range2.getFirstRow())) // left neighbor
								|| (range1.getLastColumn() + 1 == range2.getFirstColumn() && !(range1.getFirstRow() > range2.getLastRow() || range1.getLastRow() < range2.getFirstRow())) // right neighbor
								|| (range1.getFirstRow() - 1 == range2.getLastRow() && !(range1.getFirstColumn() > range2.getLastColumn() || range1.getLastColumn() < range2.getFirstColumn())) // top neighbor
								|| (range1.getLastRow() + 1 == range2.getFirstRow() && !(range1.getFirstColumn() > range2.getLastColumn() || range1.getLastColumn() < range2.getFirstColumn()))) // bottom neighbor
							throwNonRectangularAreaException(tag);
					}
				if (range1.getNumberOfCells() == 1)
					toDeleteOneCellRange.add(range1);
			}
		}
		
		if (toDeleteOneCellRange.size() != 0) {
			logger.debug("Tag \"" + tag + "\" appears in these single cells which cannot be merged: " + toDeleteOneCellRange.stream().map(range -> range.formatAsString()).collect(Collectors.joining(", ")));
			warningMessage.addMessage("Merge ranges consisting of one cell only were detected and removed from the merge instruction. See log for details.");
		}
		ret.removeAll(toDeleteOneCellRange);
		
		return ret;
	}
	
	private static void splitRange(Map<Integer, CellRangeAddress> mapColumnToCurrentlyOpenRange, List<CellRangeAddress> finalRangesList,
			CellRangeAddress range, int currentColumn, int previousRangeLastRow, int newRangeLastRow, boolean adjacentRangesAllowed, String searchedTag) {
		
		if (adjacentRangesAllowed) {

			// split vertically, meaning range is closed left of the current column and splitRange is started at the current column 
			CellRangeAddress splitRange = range.copy();
			range.setLastColumn(currentColumn - 1);
			range.setLastRow(previousRangeLastRow);
			splitRange.setFirstColumn(currentColumn);
			splitRange.setLastRow(newRangeLastRow);
			finalRangesList.add(splitRange);
			
			for (int c : mapColumnToCurrentlyOpenRange.keySet())
				if (c >= currentColumn && mapColumnToCurrentlyOpenRange.get(c) == range)
					mapColumnToCurrentlyOpenRange.put(c, splitRange);
		}
		else
			throwNonRectangularAreaException(searchedTag);
	}
	
	private static void throwNonRectangularAreaException(String tag) throws IllegalArgumentException {
		throw new IllegalArgumentException("Non-rectangular area of tag " + tag + " found in control table.");
	}
}
