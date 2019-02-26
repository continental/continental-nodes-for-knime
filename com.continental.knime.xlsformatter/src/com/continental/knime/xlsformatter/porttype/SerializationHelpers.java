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

import org.apache.commons.lang3.tuple.Triple;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

public class SerializationHelpers {

	public static void writeCellRangeAddress(CellRangeAddress value, ObjectOutput output, long masterSerialVersionUID) throws IOException {
		output.writeInt(value.getFirstRow());
		output.writeInt(value.getLastRow());
		output.writeInt(value.getFirstColumn());
		output.writeInt(value.getLastColumn());
	}
	
	public static CellRangeAddress readCellRangeAddress(ObjectInput input, long masterSerialVersionUID) throws IOException {
		int[] tempRange = new int[4];
		for (int i = 0; i < 4; i++)
			tempRange[i] = input.readInt();
		return new CellRangeAddress(tempRange[0], tempRange[1], tempRange[2], tempRange[3]); // firstRow, lastRow, firstCol, lastCol
	}
	
	
	public static void writeCellAddress(CellAddress value, ObjectOutput output, long masterSerialVersionUID) throws IOException {
		output.writeInt(value.getRow());
		output.writeInt(value.getColumn());
	}
	
	public static CellAddress readCellAddress(ObjectInput input, long masterSerialVersionUID) throws IOException {
		int row = input.readInt();
		return new CellAddress(row, input.readInt()); // row, column
	}
	
	
	public static void writeNullableInt(Integer value, ObjectOutput output, long masterSerialVersionUID) throws IOException {
		output.writeBoolean(value != null);
		if (value != null)
			output.writeInt(value);
	}
	
	public static Integer readNullableInt(ObjectInput input, long masterSerialVersionUID) throws IOException {
		if (input.readBoolean())
			return input.readInt();
		return null;
	}
	
	
	public static void writeNullableString(String value, ObjectOutput output, long masterSerialVersionUID) throws IOException {
		output.writeBoolean(value != null);
		if (value != null)
			output.writeUTF(value);
	}
	
	public static String readNullableString(ObjectInput input, long masterSerialVersionUID) throws IOException {
		if (input.readBoolean())
			return input.readUTF();
		return null;
	}
	
	
	public static void writeGroupTriple(Triple<Integer, Integer, Boolean> value, ObjectOutput output, long masterSerialVersionUID) throws IOException {
		output.writeInt(value.getLeft());
		output.writeInt(value.getMiddle());
		output.writeBoolean(value.getRight());
	}
	
	public static Triple<Integer, Integer, Boolean> readGroupTriple(ObjectInput input, long masterSerialVersionUID) throws IOException {
		int from = input.readInt();
		int to = input.readInt();
		boolean collapsed = input.readBoolean();
		return Triple.of(from, to, collapsed);
	}
	
	
	public static void writeNullableColor(Color value, ObjectOutput output, long masterSerialVersionUID) throws IOException {
		output.writeBoolean(value != null);
		if (value != null)
			output.writeInt(value.getRGB()); // incl. alpha
	}
	
	public static Color readNullableColor(ObjectInput input, long masterSerialVersionUID) throws IOException {
		if (input.readBoolean())
			return new Color(input.readInt(), true); // version with alpha as this is also written out
		return null;
	}
}
