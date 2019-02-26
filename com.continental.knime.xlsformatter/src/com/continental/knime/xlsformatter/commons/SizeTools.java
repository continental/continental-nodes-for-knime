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
import java.util.List;

import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.util.CellReference;

/**
 * Static methods helping to parse width / height instructions
 */
public class SizeTools {

	/**
	 * Parse a row/height UI field by interpreting it as a comma-separated list of a:b or a-b:c instructions
	 * @return list of pairs of 0-based row index and its desired height
	 */
	public static List<Pair<Integer, Integer>> parseRowHeightInstruction(String instruction) throws IllegalArgumentException {
		List<Pair<String, Integer>> list =  parseSizeInstructionList(instruction);
		List<Pair<Integer, Integer>> ret = new ArrayList<Pair<Integer, Integer>>();
		for (Pair<String, Integer> el : list) {
			int rowId = Commons.parseInt(el.getLeft());
			if (rowId > XlsFormatterControlTableValidator.XLS_SIZE_LIMIT_MAX_ROWS)
				throw new IllegalArgumentException("Row id " + rowId + " is out of the XLS specification.");
			ret.add(Pair.of(rowId - 1, el.getRight())); // user enters 1-based, internal is 0-based
		}
		return ret;
	}
	
	/**
	 * Parse a column/width UI field by interpreting it as a comma-separated list of a:b or a-b:c instructions
	 * @return list of pairs of 0-based column index and its desired width
	 */
	public static List<Pair<Integer, Integer>> parseColumnWidthInstructionList(String instruction) throws IllegalArgumentException {
		List<Pair<String, Integer>> list =  parseSizeInstructionList(instruction);
		List<Pair<Integer, Integer>> ret = new ArrayList<Pair<Integer, Integer>>();
		for (Pair<String, Integer> el : list) {
			int colId = -1;
			try {
				colId = CellReference.convertColStringToIndex(el.getLeft().replace("0", ""));
			} catch (Exception e) {
				throw new IllegalArgumentException("Illegal column name " + el.getLeft() + ".");
			}
			ret.add(Pair.of(colId, el.getRight()));
		}
		return ret;
	}
	
	
	private static List<Pair<String, Integer>> parseSizeInstructionList(String instruction) throws IllegalArgumentException {
		List<Pair<String, Integer>> ret = new ArrayList<Pair<String, Integer>>();
		String[] instructions = instruction.split(",");
		for (String inst : instructions) {
			inst = inst.trim();
			String[] components = inst.split(":");
			if (components.length != 2)
				throw new IllegalArgumentException("Incorrect instruction \"" + inst + "\". Must contain a colon separating object and desired height.");
			String[] idRangeParts = components[0].trim().split("-");
			if (idRangeParts.length >= 1)
				throw new IllegalArgumentException("Ranged targets (with dash) not supported yet: " + components[0].trim());
			ret.add(Pair.of(components[0], Commons.parseInt(components[1])));
		}
		return ret;
	}
}
