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

package com.continental.knime.xlsformatter.porttype;

import java.awt.Color;
import java.io.IOException;
import java.io.ObjectInput;
import java.io.ObjectOutput;
import java.util.Map;

import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

public class SerializationHelpers {

	public static void writeCellRangeAddress(CellRangeAddress value, ObjectOutput output, int serializationVersion) throws IOException {
		output.writeInt(value.getFirstRow());
		output.writeInt(value.getLastRow());
		output.writeInt(value.getFirstColumn());
		output.writeInt(value.getLastColumn());
	}
	
	public static CellRangeAddress readCellRangeAddress(ObjectInput input, int serializationVersion) throws IOException {
		int[] tempRange = new int[4];
		for (int i = 0; i < 4; i++)
			tempRange[i] = input.readInt();
		return new CellRangeAddress(tempRange[0], tempRange[1], tempRange[2], tempRange[3]); // firstRow, lastRow, firstCol, lastCol
	}
	
	
	public static void writeCellAddress(CellAddress value, ObjectOutput output, int serializationVersion) throws IOException {
		output.writeInt(value.getRow());
		output.writeInt(value.getColumn());
	}
	
	public static CellAddress readCellAddress(ObjectInput input, int serializationVersion) throws IOException {
		int row = input.readInt();
		return new CellAddress(row, input.readInt()); // row, column
	}
	
	
	public static void writeNullableInt(Integer value, ObjectOutput output, int serializationVersion) throws IOException {
		output.writeBoolean(value != null);
		if (value != null)
			output.writeInt(value);
	}
	
	public static Integer readNullableInt(ObjectInput input, int serializationVersion) throws IOException {
		if (input.readBoolean())
			return input.readInt();
		return null;
	}
	
	
	public static void writeNullableString(String value, ObjectOutput output, int serializationVersion) throws IOException {
		output.writeBoolean(value != null);
		if (value != null)
			output.writeUTF(value);
	}
	
	public static String readNullableString(ObjectInput input, int serializationVersion) throws IOException {
		if (input.readBoolean())
			return input.readUTF();
		return null;
	}
	
	
	public static void writeGroupingEntry(Integer from, Integer to, boolean isCollapsed, ObjectOutput output, int serializationVersion) throws IOException {
		output.writeInt(from);
		output.writeInt(to);
		output.writeBoolean(isCollapsed);
	}
	
	public static void addReadGroupingEntryToMap(Map<Pair<Integer, Integer>, Boolean> map, ObjectInput input, int serializationVersion) throws IOException {
		int from = input.readInt();
		int to = input.readInt();
		boolean collapsed = input.readBoolean();
		map.put(Pair.of(from, to), collapsed);
	}
	
	
	public static void writeNullableColor(Color value, ObjectOutput output, int serializationVersion) throws IOException {
		output.writeBoolean(value != null);
		if (value != null)
			output.writeInt(value.getRGB()); // incl. alpha
	}
	
	public static Color readNullableColor(ObjectInput input, int serializationVersion) throws IOException {
		if (input.readBoolean())
			return new Color(input.readInt(), true); // version with alpha as this is also written out
		return null;
	}
}
