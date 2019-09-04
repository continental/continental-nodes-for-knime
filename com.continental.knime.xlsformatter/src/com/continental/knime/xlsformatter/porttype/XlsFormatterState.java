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

import java.awt.BorderLayout;
import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.Externalizable;
import java.io.IOException;
import java.io.ObjectInput;
import java.io.ObjectInputStream;
import java.io.ObjectOutput;
import java.io.ObjectOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.swing.JComponent;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;

import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.port.PortType;
import org.knime.core.node.port.PortTypeRegistry;

import com.continental.knime.xlsformatter.apply.XlsFormatterApplyLogic;
import com.continental.knime.xlsformatter.commons.ColorTools;
import com.continental.knime.xlsformatter.commons.Commons;

// NOTE: changes to the state class (e.g. new fields) must also be reflected in XlsFormatterStateMerger

public class XlsFormatterState implements PortObject, Externalizable {

	/**
	 * A threshold value of how many cells are to be included in a String representation of a XlsFormatterState.
	 * Performance on real-life cases with thousands of formatted cells would be too low.
	 */
	public static int VIEW_CELLS_THRESHOLD = 100;
	
	public static final PortType TYPE = PortTypeRegistry.getInstance().getPortType(XlsFormatterState.class);
	public static final PortType TYPE_OPTIONAL = PortTypeRegistry.getInstance().getPortType(XlsFormatterState.class, true);
	
	/**
	 * An flag stating whether a boolean formatting option (e.g. bold or word-wrap) shall
	 * * not be modified,
	 * * explicitly be turned off, or
	 * * explicitly be turned on.
	 */
	public enum FormattingFlag { UNMODIFIED, OFF, ON }
	
	/**
	 * Formatting options of a cell's horizontal alignment.
	 */
	public enum CellAlignmentHorizontal { UNMODIFIED, LEFT, CENTER, RIGHT, JUSTIFY }
	
	/**
	 * Formatting options of a cell's vertical alignment.
	 */
	public enum CellAlignmentVertical { UNMODIFIED, TOP, MIDDLE, BOTTOM }
	
	/**
	 * Formatting style options of a border's edge.
	 */
	public enum BorderStyle { UNMODIFIED, NONE, NORMAL, DOUBLE, DASHED, THICK, DASHED_THICK, EXTRA_THICK }
	
	/**
	 * Formatting options of a cell's background.
	 */
	public enum FillPattern { UNMODIFIED, NONE, SOLID_BACKGROUND_COLOR, DIAGONAL, HORIZONTAL, VERTICAL, DOTTED }
	
	/**
	 * Data type conversion from a KNIME String cell to a POI String, Numeric, or Boolean cell (with automatic POI style
	 * definition in case of date/time).
	 */
	public enum CellDataType {
		UNMODIFIED,
		NUMERIC,
		BOOLEAN,
		FORMULA,
		LOCALDATE,
		LOCALDATETIME,
		LOCALTIME;
		
		@Override
		public String toString() {
			switch (this) {
			case LOCALDATE:
				return "local date";
			case LOCALDATETIME:
				return "local date/time";
			case LOCALTIME:
				return "local time";
			default:
				return super.toString().toLowerCase();
			}
		}
		
		/**
		 * Gets a default text format for the date/time typed options.
		 */
		public String getDateTextFormat() {
			switch (this) {
			case LOCALDATE:
				return "yyyy-MM-dd";
			case LOCALDATETIME:
				return "yyyy-MM-dd hh:mm:ss";
			case LOCALTIME:
				return "hh:mm:ss";
			default:
				throw new IllegalArgumentException("getDateTextFormat() can be called on date/time types only.");
			}
		}
		
		/**
		 * Gets a String array of allowed incoming text formats per date/time typed option.
		 * @return
		 */
		public String[] getDateTextFormats() {
			switch (this) {
			case LOCALDATE:
				return new String[] { "yyyy-MM-dd" };
			case LOCALDATETIME:
				return new String[] { "yyyy-MM-dd'T'hh:mm:ss", "yyyy-MM-dd'T'hh:mm", "yyyy-MM-dd hh:mm:ss" };
			case LOCALTIME:
				return new String[] { "hh:mm", "hh:mm:ss" };
			default:
				throw new IllegalArgumentException("getDateTextFormats() can be called on date/time types only.");
			}
		}
	}
	
	/**
	 * An instruction set for adding a cell comment.
	 */
	public static class Comment {
		public String author;
		public String text;
		
		public static Comment readFromExternal(ObjectInput input, int serializationVersion) throws IOException, ClassNotFoundException {
			Comment ret = new Comment();
			input.readInt(); // placeholder for future type implementations
			ret.author = SerializationHelpers.readNullableString(input, serializationVersion);
			ret.text = SerializationHelpers.readNullableString(input, serializationVersion);
			return ret;
		}

		public void writeExternal(ObjectOutput output, int serializationVersion) throws IOException {
			output.writeInt(0); // placeholder for future type implementations (0 being default)
			SerializationHelpers.writeNullableString(author, output, serializationVersion);
			SerializationHelpers.writeNullableString(text, output, serializationVersion);
		}
	}
	
	/**
	 * An instruction set for conditional formatting, which is re-used over multiple cells by associating it to cell ranges.
	 */
	public static class ConditionalFormattingSet {
		
		/**
		 * List of background color formatting steps (pairs) of numeric value and color
		 */
		public List<Pair<Double, Color>> backgroundScaleFixpoints = new ArrayList<Pair<Double, Color>>();
		
		/**
		 * Get a String representation of this ConditionalFormattingSet. Two different ConditionalFormattingSet must
		 * yield two different Strings returned by this method.
		 */
		public String toString() {
			StringBuilder sb = new StringBuilder();
			sb.append("cf:");
			if (backgroundScaleFixpoints == null || backgroundScaleFixpoints.size() == 0)
				sb.append("-");
			else
				for (Pair<Double, Color> pair : backgroundScaleFixpoints)
					sb.append(pair.getLeft() + ":" + ColorTools.colorToXlsfColorString(pair.getRight()) + ";");
			return sb.toString();
		}
		
		public static ConditionalFormattingSet readFromExternal(ObjectInput input, int serializationVersion) throws IOException, ClassNotFoundException {
			int tempSize = input.readInt();
			ConditionalFormattingSet ret = new ConditionalFormattingSet();
			ret.backgroundScaleFixpoints = new ArrayList<Pair<Double, Color>>(tempSize);
			for (int i = 0; i < tempSize; i++) {
				double step = input.readDouble();
				Color color = SerializationHelpers.readNullableColor(input, serializationVersion);
				ret.backgroundScaleFixpoints.add(Pair.of(step, color));
			}
			return ret;
		}

		public void writeExternal(ObjectOutput output, int serializationVersion) throws IOException {
			output.writeInt(backgroundScaleFixpoints.size());
			for (Pair<Double, Color> pair : backgroundScaleFixpoints) {
				output.writeDouble(pair.getLeft());
				SerializationHelpers.writeNullableColor(pair.getRight(), output, serializationVersion);
			}
		}
	}
	
	/**
	 * Formatting instructions of one edge of one cell's border.
	 * Style and color are independent of each other, meaning that even with unmodified style
	 * but defined color, the apply logic will implement this change in POI.
	 */
	public static class BorderEdge {
		
		public BorderStyle style;
		public Color color = null;
		
		public BorderEdge(BorderStyle style, Color color) {
			this.style = style;
			this.color = color;
		}
		
		public BorderEdge(BorderStyle style) {
			this.style = style;
		}
		
		/**
		 * Get a String representation of this BorderEdge. Two different BorderEdge must
		 * yield two different Strings returned by this method.
		 */
		public String toString() {
			return resolveEnumUnmodified(style.toString()) + "," + (color == null ? "-" : ColorTools.colorToXlsfColorString(color)) + ";";
		}
		
		/**
		 * Overwrites this edge with values of the new one, except they are marked as UNMODIFIED or null. 
		 * @param overwritingEdge The incoming (newer) edge that shall overwrite this edge.
		 */
		public void mergeIn (BorderEdge overwritingEdge) {
			if (overwritingEdge.style != BorderStyle.UNMODIFIED)
				style = overwritingEdge.style;
			if (overwritingEdge.color != null)
				color = overwritingEdge.color;
		}
		
		/**
		 * Merges two border edges. In case of conflicting information, master wins.
		 * No cloning is performed, so make sure that the slave's color object can
		 * safely be pointed to from within master.
		 */
		public static void merge(BorderEdge master, BorderEdge slave) {
			if (master != null && slave != null) {
				if (master.style == BorderStyle.UNMODIFIED)
					master.style = slave.style;
				if (master.color == null)
					master.color = slave.color;
			}
		}
		
		public static BorderEdge readFromExternal(ObjectInput input, int serializationVersion) throws IOException, ClassNotFoundException {
			BorderStyle style;
			switch (input.readByte()) {
				case 1:
					style = BorderStyle.NONE;
					break;	
				case 2:
					style = BorderStyle.NORMAL;
					break;
				case 3:
					style = BorderStyle.THICK;
					break;
				case 4:
					style = BorderStyle.EXTRA_THICK;
					break;
				case 5:
					style = BorderStyle.DOUBLE;
					break;
				case 6:
					style = BorderStyle.DASHED;
					break;
				case 7:
					style = BorderStyle.DASHED_THICK;
					break;
				case 0:
				default:
					style = BorderStyle.UNMODIFIED;
					break;
			}
			Color color = SerializationHelpers.readNullableColor(input, serializationVersion);
			return new BorderEdge(style, color);
		}

		public void writeExternal(ObjectOutput output, int serializationVersion) throws IOException {
			switch (style) {
				case NONE:
					output.writeByte(1);
					break;
				case NORMAL:
					output.writeByte(2);
					break;
				case THICK:
					output.writeByte(3);
					break;
				case EXTRA_THICK:
					output.writeByte(4);
					break;
				case DOUBLE:
					output.writeByte(5);
					break;
				case DASHED:
					output.writeByte(6);
					break;
				case DASHED_THICK:
					output.writeByte(7);
					break;
				case UNMODIFIED:
				default:
					output.writeByte(0);
			}
			SerializationHelpers.writeNullableColor(color, output, serializationVersion);
		}
	}
	
	
	
	
	
	
	/**
	 * The formatting state of a single cell.
	 */
	public static class CellState {
		
		public Integer fontSize = null;
		public FormattingFlag fontBold = FormattingFlag.UNMODIFIED;
		public FormattingFlag fontItalic = FormattingFlag.UNMODIFIED;
		public FormattingFlag fontUnderline = FormattingFlag.UNMODIFIED;
		public Color fontColor = null;
		
		public CellAlignmentHorizontal cellHorizontalAlignment = CellAlignmentHorizontal.UNMODIFIED;
		public CellAlignmentVertical cellVerticalAlignment = CellAlignmentVertical.UNMODIFIED;
		public FormattingFlag wrapText = FormattingFlag.UNMODIFIED;
		public Integer textTiltDegree = null;
		
		/**
		 * A cell's background color. Note that when not combined with a fillPattern, this in POI would be called a foregroundColor. 
		 */
		public Color backgroundColor = null;
		
		/**
		 *  A special pattern for the cell. Like in POI, setting a backgroundColor requires a fillPattern of SOLID_BACKGROUND_COLOR.
		 */
		public FillPattern fillPattern = FillPattern.UNMODIFIED;  
		
		/**
		 * The foreground color of a special fill pattern (e.g. diagonals).
		 * Unlike in POI, a solid background color is handled via the backgroundColor property, not this fillForegroundColor.
		 */
		public Color fillForegroundColor = null;
		
		/**
		 * A conditional formatting instruction for this cell.
		 * Note that this is irrelevant for unique style derivation as in POI it's applied on sheet/range level, not on cell level.
		 */
		public ConditionalFormattingSet conditionalFormat = null;
		
		/**
		 * The cell data type that shall be set by POI. This is not directly a formatting instruction, but a underlying data type.
		 * Entries other than UNMODIFIED represent a conversion from a String cell to some other data type.  
		 */
		public CellDataType cellDataType = CellDataType.UNMODIFIED;
		
		/**
		 * The text format that controls how (esp. numeric) values are displayed.
		 */
		public String textFormat = null;
		
		public BorderEdge borderTop = null;
		public BorderEdge borderBottom = null;
		public BorderEdge borderLeft = null;
		public BorderEdge borderRight = null;
		public BorderEdge borderDiagonalSlash = null; // not currently implemented, included in serialization for future proofing reasons
		public BorderEdge borderDiagonalBackslash = null; // not currently implemented, included in serialization for future proofing reasons
		
		/**
		 * A hyperlink that shall be set for a cell.
		 * Note that hyperlink is irrelevant for unique POI style derivation.
		 */
		public String hyperlink = null;
		
		/**
		 * A comment that shall be added to a cell.
		 * Note that comments are irrelevant for unique POI style derivation.
		 */
		public Comment comment = null;
		
		/**
		 * Transform this cell's font specification to String in order to later on match equal fonts to generate only a joint Font object in POI
		 */
		public String fontDefinitionToShortString() {
			StringBuilder sb = new StringBuilder();
			sb.append("s:" + (fontSize == null ? "-," : fontSize + ","));
			sb.append("b:" + resolveEnumUnmodified(fontBold.toString()) + ",");
			sb.append("i:" + resolveEnumUnmodified(fontItalic.toString()) + ",");
			sb.append("u:" + resolveEnumUnmodified(fontUnderline.toString()) + ",");
			sb.append("c:" + (fontColor == null ? "-" : ColorTools.colorToXlsfColorString(fontColor)));
			return sb.toString();
		}
		
		/**
		 * Transform this cell's format specification to String in order to later on match equal styles and to generate only a joint CellStyle object in POI.
		 * Makes use of shared Font specification internally already.
		 * This function is only used during computation and its results are never serialized. Hence it may differ between versions.
		 * @param restrictToPoiStyleRelevantProperties If true, only information is included which leads to a POI style being necessary. If false, all cell state information is included. 
		 * @param includeBorderFormatting If false, border formatting is excluded (typical use case is around cell merging functions, where border information can be different per cell of a merged range, all other properties should be the same)
		 */
		public String cellFormatToShortString(boolean restrictToPoiStyleRelevantProperties, boolean includeBorderFormatting) {
			StringBuilder sb = new StringBuilder();
			sb.append("font:");
			sb.append(fontDefinitionToShortString());
			sb.append(";alignH:");
			sb.append(resolveEnumUnmodified(cellHorizontalAlignment.toString()));
			sb.append(";alignV:");
			sb.append(resolveEnumUnmodified(cellVerticalAlignment.toString()));
			sb.append(";wrap:");
			sb.append(resolveEnumUnmodified(wrapText.toString()));
			sb.append(";tilt:");
			sb.append(textTiltDegree == null ? "-" : textTiltDegree);
			sb.append(";bgCol:");
			sb.append(backgroundColor == null ? "-" : ColorTools.colorToXlsfColorString(backgroundColor));
			sb.append(";fillP:");
			sb.append(resolveEnumUnmodified(fillPattern.toString()));
			sb.append(";patCol:");
			sb.append(fillForegroundColor == null ? "-" : ColorTools.colorToXlsfColorString(fillForegroundColor));
			sb.append(";dt:");
			sb.append(restrictToPoiStyleRelevantProperties && (cellDataType == CellDataType.NUMERIC || cellDataType == CellDataType.BOOLEAN) ?
					"/" : resolveEnumUnmodified(cellDataType.toString()));
			sb.append(";nf:");
			sb.append(textFormat == null ? "-" : textFormat);
			if (includeBorderFormatting) {
				sb.append(";borders:T:");
				sb.append(borderTop == null ? "-;" : borderTop.toString());
				sb.append("B:");
				sb.append(borderBottom == null ? "-;" : borderBottom.toString());
				sb.append("L:");
				sb.append(borderLeft == null ? "-;" : borderLeft.toString());
				sb.append("R:");
				sb.append(borderRight == null ? "-;" : borderRight.toString());
			}
			if (!restrictToPoiStyleRelevantProperties) {
				sb.append(conditionalFormat == null ? "cf:-" : conditionalFormat.toString());
				sb.append(";cmnt:");
				sb.append(comment == null ? "-" : "\"" + comment.author + "\":\"" + comment.text + "\"");
				sb.append(";hl:");
				sb.append(hyperlink == null ? "-" : "\"" + hyperlink + "\"");
			}
			return sb.toString();
		}
		
		private static String nonFormattingStateString = null;
		
		/**
		 * Returns the cellFormatToShortString() return code on an "empty" CellState that doesn't contain any formatting
		 * instructions that would require a POI style, but maybe just a hyperlink.
		 */
		public static String getNonFormattingStateString() {
			if (nonFormattingStateString == null)
				nonFormattingStateString = (new CellState()).cellFormatToShortString(true, true); 
			return nonFormattingStateString;
		}
		
		private static String defaultFontShortString = null;
		
		/**
		 * Returns the fontDefinitionToShortString() return code on an "empty" CellState with hence a default font.
		 */
		public static String getDefaultFontShortString() {
			if (defaultFontShortString == null)
				defaultFontShortString = (new CellState()).fontDefinitionToShortString();
			return defaultFontShortString;
		}
		
		private static FormattingFlag getFormattingFlagFromSerializedByte(byte value) {
			switch (value) {
			case 1:
				return FormattingFlag.OFF;
			case 2:
				return FormattingFlag.ON;
			case 0:
			default:
				return FormattingFlag.UNMODIFIED;
			}
		}

		public static CellState readFromExternal(ObjectInput input, int serializationVersion) throws IOException, ClassNotFoundException {
			CellState ret = new CellState();
			ret.fontSize = SerializationHelpers.readNullableInt(input, serializationVersion);
			ret.fontBold = getFormattingFlagFromSerializedByte(input.readByte());
			ret.fontItalic = getFormattingFlagFromSerializedByte(input.readByte());
			ret.fontUnderline = getFormattingFlagFromSerializedByte(input.readByte());
			ret.fontColor = SerializationHelpers.readNullableColor(input, serializationVersion);
			ret.wrapText = getFormattingFlagFromSerializedByte(input.readByte());
			byte cellHorizontalAlignmentByte = input.readByte();
			switch (cellHorizontalAlignmentByte) {
				case 1:
					ret.cellHorizontalAlignment = CellAlignmentHorizontal.LEFT;
					break;
				case 2:
					ret.cellHorizontalAlignment = CellAlignmentHorizontal.RIGHT;
					break;
				case 3:
					ret.cellHorizontalAlignment = CellAlignmentHorizontal.CENTER;
					break;
				case 4:
					ret.cellHorizontalAlignment = CellAlignmentHorizontal.JUSTIFY;
					break;
				case 0:
				default:
					ret.cellHorizontalAlignment = CellAlignmentHorizontal.UNMODIFIED;
					break;
			}
			byte cellVerticalAlignmentByte = input.readByte();
			switch (cellVerticalAlignmentByte) {
			case 1:
				ret.cellVerticalAlignment = CellAlignmentVertical.TOP;
				break;
			case 2:
				ret.cellVerticalAlignment = CellAlignmentVertical.MIDDLE;
				break;
			case 3:
				ret.cellVerticalAlignment = CellAlignmentVertical.BOTTOM;
				break;
			case 0:
			default:
				ret.cellVerticalAlignment = CellAlignmentVertical.UNMODIFIED;
				break;
			}
			byte fillPattern = input.readByte();
			switch (fillPattern) {
			case 1:
				ret.fillPattern = FillPattern.NONE;
				break;
			case 2:
				ret.fillPattern = FillPattern.DIAGONAL;
				break;
			case 3:
				ret.fillPattern = FillPattern.HORIZONTAL;
				break;
			case 4:
				ret.fillPattern = FillPattern.VERTICAL;
				break;
			case 5:
				ret.fillPattern = FillPattern.DOTTED;
				break;
			case 10:
				ret.fillPattern = FillPattern.SOLID_BACKGROUND_COLOR;
				break;
			case 0:
			default:
				ret.fillPattern = FillPattern.UNMODIFIED;
				break;
			}
			ret.textTiltDegree = SerializationHelpers.readNullableInt(input, serializationVersion);
			ret.backgroundColor = SerializationHelpers.readNullableColor(input, serializationVersion);
			ret.fillForegroundColor = SerializationHelpers.readNullableColor(input, serializationVersion);
			ret.textFormat = SerializationHelpers.readNullableString(input, serializationVersion);
			ret.hyperlink = SerializationHelpers.readNullableString(input, serializationVersion);
			input.readByte(); // placeholder for future hyperlink type implementation, 0 meaning DEFAULT
			if (serializationVersion >= 2)
				if (input.readBoolean())
					ret.comment = Comment.readFromExternal(input, serializationVersion);
			if (input.readBoolean())
				ret.conditionalFormat = ConditionalFormattingSet.readFromExternal(input, serializationVersion);
			if (input.readBoolean())
				ret.borderTop = BorderEdge.readFromExternal(input, serializationVersion);
			if (input.readBoolean())
				ret.borderBottom = BorderEdge.readFromExternal(input, serializationVersion);
			if (input.readBoolean())
				ret.borderLeft = BorderEdge.readFromExternal(input, serializationVersion);
			if (input.readBoolean())
				ret.borderRight = BorderEdge.readFromExternal(input, serializationVersion);
			if (input.readBoolean())
				ret.borderDiagonalSlash = BorderEdge.readFromExternal(input, serializationVersion);
			if (input.readBoolean())
				ret.borderDiagonalBackslash = BorderEdge.readFromExternal(input, serializationVersion);
			byte dataType = input.readByte();
			switch (dataType) {
			case 1:
				ret.cellDataType = CellDataType.NUMERIC;
				break;
			case 2:
				ret.cellDataType = CellDataType.BOOLEAN;
				break;
			case 3:
				ret.cellDataType = CellDataType.FORMULA;
				break;
			case 4:
				ret.cellDataType = CellDataType.LOCALDATE;
				break;
			case 5:
				ret.cellDataType = CellDataType.LOCALDATETIME;
				break;
			case 6:
				ret.cellDataType = CellDataType.LOCALTIME;
				break;
			case 0:
			default:
				ret.cellDataType = CellDataType.UNMODIFIED;
				break;
			}
			return ret;
		}
		
		/**
		 * Externalized writing method.
		 */
		public void writeExternal(ObjectOutput output, int serializationVersion) throws IOException {
			SerializationHelpers.writeNullableInt(fontSize, output, serializationVersion);
			output.writeByte(fontBold == FormattingFlag.UNMODIFIED ? 0 : (fontBold == FormattingFlag.OFF ? 1 : 2));
			output.writeByte(fontItalic == FormattingFlag.UNMODIFIED ? 0 : (fontItalic == FormattingFlag.OFF ? 1 : 2));
			output.writeByte(fontUnderline == FormattingFlag.UNMODIFIED ? 0 : (fontUnderline == FormattingFlag.OFF ? 1 : 2));
			SerializationHelpers.writeNullableColor(fontColor, output, serializationVersion);
			output.writeByte(wrapText == FormattingFlag.UNMODIFIED ? 0 : (wrapText == FormattingFlag.OFF ? 1 : 2));
			switch (cellHorizontalAlignment) {
				case LEFT:
					output.writeByte(1);
					break;
				case RIGHT:
					output.writeByte(2);
					break;
				case CENTER:
					output.writeByte(3);
					break;
				case JUSTIFY:
					output.writeByte(4);
					break;
				case UNMODIFIED:
				default:
					output.writeByte(0);
					break;
			}
			switch (cellVerticalAlignment) {
				case TOP:
					output.writeByte(1);
					break;
				case MIDDLE:
					output.writeByte(2);
					break;
				case BOTTOM:
					output.writeByte(3);
					break;
				case UNMODIFIED:
				default:
					output.writeByte(0);
					break;
			}
			switch (fillPattern) { 
			case NONE:
				output.writeByte(1);
				break;
			case DIAGONAL:
				output.writeByte(2);
				break;
			case HORIZONTAL:
				output.writeByte(3);
				break;
			case VERTICAL:
				output.writeByte(4);
				break;
			case DOTTED:
				output.writeByte(5);
				break;
			case SOLID_BACKGROUND_COLOR:
				output.writeByte(10);
				break;
			case UNMODIFIED:
			default:
				output.writeByte(0);
				break;
			}
			SerializationHelpers.writeNullableInt(textTiltDegree, output, serializationVersion);
			SerializationHelpers.writeNullableColor(backgroundColor, output, serializationVersion);
			SerializationHelpers.writeNullableColor(fillForegroundColor, output, serializationVersion);
			SerializationHelpers.writeNullableString(textFormat, output, serializationVersion);
			
			SerializationHelpers.writeNullableString(hyperlink, output, serializationVersion);
			output.writeByte(0); // placeholder for future hyperlink type implementation, 0 meaning DEFAULT
			
			output.writeBoolean(comment != null);
			if (comment != null)
				comment.writeExternal(output, serializationVersion);
			
			output.writeBoolean(conditionalFormat != null);
			if (conditionalFormat != null)
				conditionalFormat.writeExternal(output, serializationVersion);
			
			for (BorderEdge border : new BorderEdge[] { borderTop, borderBottom, borderLeft, borderRight, borderDiagonalSlash, borderDiagonalBackslash }) {
				output.writeBoolean(border != null);
				if (border != null)
					border.writeExternal(output, serializationVersion);
			}
			
			switch (cellDataType) {
			case NUMERIC:
				output.writeByte(1);
				break;
			case BOOLEAN:
				output.writeByte(2);
				break;
			case FORMULA:
				output.writeByte(3);
				break;
			case LOCALDATE:
				output.writeByte(4);
				break;
			case LOCALDATETIME:
				output.writeByte(5);
				break;
			case LOCALTIME:
				output.writeByte(6);
				break;
			case UNMODIFIED:
			default:
				output.writeByte(0);
				break;
			}
		}
	}
	
	
	
	/**
	 * The serialization UID that must remain constant over releases for the Java Serialization to
	 * still recognize the same class.
	 */
	private static final long serialVersionUID = 1L;
	
	/**
	 * The serialization version controlling backward compatibility for future releases.
	 * It is used as the "one and only" master serial version, even for subclasses.
	 */
	private static final int masterSerializationVersion = 2; // see a history of versions in the comment to the writeExternal method below
	
	/**
	 * The state of a specific sheet (incl. the CellStates of its cells)
	 */
	public class SheetState {
		
		
		/**
		 * Map of cell address to its internal cell state, holding all cell relevant formatting + additional instructions.
		 */
		public Map<CellAddress, CellState> cells = new HashMap<CellAddress, CellState>();
		
		/**
		 * Cell at whose top left corner the sheet shall be frozen.
		 */
		public CellAddress freezeSheetAtTopLeftCornerOfCell = null;
		
		/**
		 * Cell range that an auto filter shall be applied on (incl. the data part, not only the table header).
		 */
		public CellRangeAddress autoFilterRange = null;
		
		/**
		 * Map of column index to column width. Null in value means auto-sized column.
		 */
		public Map<Integer, Double> columnWidths = new HashMap<Integer, Double>();
		
		/**
		 * Map of row index to row height.
		 */
		public Map<Integer, Double> rowHeights = new HashMap<Integer, Double>();
		
		/**
		 * Set of row IDs to hide.
		 */
		public Set<Integer> hiddenRows = new HashSet<Integer>();
		
		/**
		 * Set of column IDs to hide.
		 */
		public Set<Integer> hiddenColumns = new HashSet<Integer>();
		
		/**
		 * Cell ranges to be merged.
		 */
		public List<CellRangeAddress> mergeRanges = new ArrayList<CellRangeAddress>();
		
		/**
		 * Map of column groups: <<fromCol, toCol>, isCollapsed>
		 */
		public Map<Pair<Integer, Integer>, Boolean> columnGroups = new HashMap<Pair<Integer, Integer>, Boolean>();
		
		/**
		 * Map of row groups: <<fromRow, toRow>, isCollapsed>
		 */
		public Map<Pair<Integer, Integer>, Boolean> rowGroups = new HashMap<Pair<Integer, Integer>, Boolean>();
	
		/**
		 * Checks whether this state is empty, meaning all data instruction-storing structures are empty.
		 */
		public boolean isEmpty() {
			return cells.size() == 0 && freezeSheetAtTopLeftCornerOfCell == null && autoFilterRange == null &&
					rowHeights.size() == 0 && columnWidths.size() == 0 && hiddenRows.size() == 0 && hiddenColumns.size() == 0 &&
					mergeRanges.size() == 0 && columnGroups.size() == 0 && rowGroups.size() == 0;
		}
	}
	
	/**
	 * A map relating a sheet name to its SheetState.
	 * A key of null represents the default sheet (with POI index 0).
	 */
	public Map<String, SheetState> sheetStates = new HashMap<String, SheetState>();
	
	/**
	 * The explicitly (or, in case of null, implicitly) selected sheet to be edited in case of formatting nodes adding instructions.
	 * Only valid in combination with analyzing the sheetStates map (size >= 2 invalidates this field, as then nothing may be added anymore).
	 */
	private String designatedSheetnameForModifications = null;
	
	/**
	 * XlsFormatterState constructor without any action as all relevant fields are instantiated.
	 */
	public XlsFormatterState() { }
	
	/**
	 * Checks whether this sheet state is empty, meaning all data instruction-storing structures are empty.
	 */
	public boolean isEmpty() {
		return sheetStates.size() == 0 || !sheetStates.values().stream().filter(s -> !s.isEmpty()).findAny().isPresent(); 
	}
	
	/**
	 * Returns the sheet state that any additional formatting instructions can be added to.
	 * Throws an exception if this state is closed for editing (e.g. due to containing >= 2 sheet states).
	 */
	public SheetState getCurrentSheetStateForModification() throws Exception {
		
		// check whether no sheet definition/selection node has yet prepared a SheetState and do so for the default case here
		if (sheetStates.size() == 0)
			sheetStates.put(null, new SheetState());
		
		// check whether this state contains already multiple sheet states and is hence closed for editing
		if (sheetStates.size() >= 2)
			throw new Exception("This XLS Formatting Port already addresses multiple sheets. Adding further instructions is hence not possible, as it would be undefined which sheet they target.");
		
		return sheetStates.get(designatedSheetnameForModifications);
	}
	
	/**
	 * Sets the sheet that subsequent formatting instructions shall be targeted to.
	 * Ensures correct internal state (i.e. creating the corresponding map entry of name and a new SheetState).
	 */
	public void setCurrentSheetForModification(String sheetName) throws Exception {
		designatedSheetnameForModifications = sheetName;
		if (sheetStates.containsKey(sheetName))
			throw new Exception("Coding issue: The XLS Formatting port shall be pointed to specific sheets only for fresh instances.");
		sheetStates.put(sheetName, new SheetState());
	}
	
	@Override
	public String getSummary() {
		if (isEmpty())
			return "XLS Formatting port object, yet without formatting instructions";
		
		int sheetLevelInstructions = sheetStates.values().stream().mapToInt(s -> {
			return
				(s.freezeSheetAtTopLeftCornerOfCell == null ? 0 : 1) +
				(s.autoFilterRange == null ? 0 : 1) +
				s.columnWidths.size() +
				s.rowHeights.size() +
				s.hiddenRows.size() +
				s.hiddenColumns.size() +
				s.mergeRanges.size() +
				s.columnGroups.size() +
				s.rowGroups.size();
			}).sum();
		return "XLS Formatting port object: " + Commons.resolvePluralString((int)sheetStates.values().stream().filter(s -> !s.isEmpty()).count(), "sheet(s)") + ", " + 
				Commons.resolvePluralString(sheetStates.values().stream().mapToInt(s -> s.cells.size()).sum(), "cell(s) with instructions") + ", " + 
				Commons.resolvePluralString(sheetLevelInstructions, "sheet-level instruction(s)");
	}
	
	/**
	 * Gets a full specification, e.g. for viewing the port state in the UI
	 */
	public String toLongString(boolean cutLongText) {
		StringBuilder sb = new StringBuilder();
		for (String sheetName : sheetStates.keySet()) {
			sb.append("##### sheet ");
			sb.append(sheetName == null ? "[default]" : "\"" + sheetName + "\"");
			sb.append(" #####");
			SheetState state = sheetStates.get(sheetName);
			sb.append("\nfreezeSheetAtTopLeftCornerOfCell: ");
			sb.append(state.freezeSheetAtTopLeftCornerOfCell == null ? "-" : state.freezeSheetAtTopLeftCornerOfCell.toString());
			sb.append("\nautoFilterRange: ");
			sb.append(state.autoFilterRange == null ? "-" : state.autoFilterRange.formatAsString());
			sb.append("\nrowHeights: ");
			for (int i : state.rowHeights.keySet())
				sb.append(i + ":" +  state.rowHeights.get(i));
			sb.append("\ncolumnWidths: ");
			for (int i : state.columnWidths.keySet())
				sb.append(CellReference.convertNumToColString(i) + ":" + state.columnWidths.get(i) + " ");
			sb.append("\nhiddenRows: ");
			for (int i : state.hiddenRows)
				sb.append(i + " ");
			sb.append("\nhiddenColumns: ");
			for (int i : state.hiddenColumns)
				sb.append(i + " ");
			sb.append("\nmergeRanges: ");
			for (CellRangeAddress range : state.mergeRanges)
				sb.append(range.formatAsString() + " ");
			sb.append("\ncolumnGroups: ");
			for (Map.Entry<Pair<Integer, Integer>, Boolean> group : state.columnGroups.entrySet())
				sb.append(CellReference.convertNumToColString(group.getKey().getLeft()) + ":" + CellReference.convertNumToColString(group.getKey().getRight()) + "," + (group.getValue() ? "collapsed" : "opened") + " ");
			sb.append("\nrowGroups: ");
			for (Map.Entry<Pair<Integer, Integer>, Boolean> group : state.rowGroups.entrySet())
				sb.append(group.getKey().getLeft() + ":" + group.getKey().getRight() + "," + (group.getValue() ? "collapsed" : "opened") + " ");
			sb.append("\n");
			int iteration = 0;
			for (CellAddress address : state.cells.keySet()) {
				sb.append("\n" + address.formatAsString() + ": ");
				sb.append(state.cells.get(address).cellFormatToShortString(false, true));
				if (iteration++ >= VIEW_CELLS_THRESHOLD && cutLongText) {
					sb.append("\n[...], " + state.cells.size() + " total cells with instructions");
					break;
				}
			}
			sb.append("\n\n");
		}
		sb.append("\n");
		sb.append(XlsFormatterApplyLogic.getDerivedStyleComplexityMessage(this, null, null));
		return sb.toString();
	}

	@Override
	public void readExternal(ObjectInput input) throws IOException, ClassNotFoundException {
		int tempSize;
		int temp;
		int readSerialVersion = (int)input.readLong();
		
		// check for cases of unsupported upward compatibility (see explanation below in the writeExternal method comment)
		int earliestSerializationVersionCapableOfReadingThis = 1; // default 1 because it equals the serialVersionId 1 of the code state that did not write this field itself yet (doesn't matter though since an upward scenario does not exist for v1 not having this logic inbuilt at all)
		if (readSerialVersion >= 2) // version 1 did not yet have this element
			earliestSerializationVersionCapableOfReadingThis = input.readInt();
		
		if (masterSerializationVersion < readSerialVersion && masterSerializationVersion < earliestSerializationVersionCapableOfReadingThis)
			throw new ClassNotFoundException("You are trying to read a XLS Formatting state that has been written with a newer version of the extension than the one currently executing. Upward compatibility is not supported. Please update to the newest version.");
		
		// read sheet header and loop all sheets
		int numberOfSheets = 1; // default for v1 is one sheet of name null
		if (readSerialVersion >= 2) // version 1 did not yet have this element
			numberOfSheets = input.readInt();
		
		for (int iSheet = 0; iSheet < numberOfSheets; iSheet++) {
			
			// handle sheet logic of this loop iteration:
			SheetState sheetState = new SheetState();
			String sheetName = null;
			if (readSerialVersion >= 2) // version 1 did not yet have this element
				sheetName = SerializationHelpers.readNullableString(input, readSerialVersion);
			if (sheetStates.containsKey(sheetName))
				throw new IOException("Coding issue: Invalid persisted XLS Formatting state, sheet names must be unique");
			sheetStates.put(sheetName, sheetState);
			
			// read cells map:
			tempSize = input.readInt();
			sheetState.cells = new HashMap<CellAddress, CellState>(tempSize);
			for (int i = 0; i < tempSize; i++) {
				CellAddress address = SerializationHelpers.readCellAddress(input, readSerialVersion);
				sheetState.cells.put(
						address,
						CellState.readFromExternal(input, readSerialVersion));
			}
			
			if (input.readBoolean())
				sheetState.freezeSheetAtTopLeftCornerOfCell = SerializationHelpers.readCellAddress(input, readSerialVersion);
			
			if (input.readBoolean())
				sheetState.autoFilterRange = SerializationHelpers.readCellRangeAddress(input, readSerialVersion);
			
			tempSize = input.readInt();
			sheetState.rowHeights = new HashMap<Integer, Double>(tempSize);
			for (int i = 0; i < tempSize; i++) {
				temp = input.readInt();
				sheetState.rowHeights.put(temp, input.readDouble());
			}
			
			tempSize = input.readInt();
			sheetState.columnWidths = new HashMap<Integer, Double>(tempSize);
			for (int i = 0; i < tempSize; i++) {
				temp = input.readInt(); // key, i.e. column index
				boolean autoSizeColumn = !input.readBoolean();
				Double columnWidth = null;
				if (!autoSizeColumn)
					columnWidth = input.readDouble();
				sheetState.columnWidths.put(temp, columnWidth);
			}
			
			tempSize = input.readInt();
			sheetState.hiddenRows = new HashSet<Integer>(tempSize);
			for (int i = 0; i < tempSize; i++)
				sheetState.hiddenRows.add(input.readInt());
			
			tempSize = input.readInt();
			sheetState.hiddenColumns = new HashSet<Integer>(tempSize);
			for (int i = 0; i < tempSize; i++)
				sheetState.hiddenColumns.add(input.readInt());
			
			tempSize = input.readInt();
			sheetState.mergeRanges = new ArrayList<CellRangeAddress>(tempSize);
			for (int i = 0; i < tempSize; i++)
				sheetState.mergeRanges.add(SerializationHelpers.readCellRangeAddress(input, readSerialVersion));
			
			tempSize = input.readInt();
			sheetState.columnGroups = new HashMap<Pair<Integer, Integer>, Boolean>();
			for (int i = 0; i < tempSize; i++)
				SerializationHelpers.addReadGroupingEntryToMap(sheetState.columnGroups, input, readSerialVersion);
			
			tempSize = input.readInt();
			sheetState.rowGroups = new HashMap<Pair<Integer, Integer>, Boolean>();
			for (int i = 0; i < tempSize; i++)
				SerializationHelpers.addReadGroupingEntryToMap(sheetState.rowGroups, input, readSerialVersion);
		}
		
		// also update the internal state for next sheet to modify (esp. as the serialization logic
		// is used to clone objects in memory)
		if (numberOfSheets == 1)
			designatedSheetnameForModifications = sheetStates.keySet().iterator().next();
	}

	
	
	@Override
	public void writeExternal(ObjectOutput output) throws IOException {
		
		// write the internal serialization version (code state) based on which even future code can detect how
		// to deserialize this object stream (for downward compatibility)
		output.writeLong((long)masterSerializationVersion);
		
		/* Earliest old serialization version that this file can still be read/deserialized with.
		 * This concept provides a chance to selectively allow upward compatibility. (Downward compatibility shall always be
		 * supported.) When reading a file written by this method in an older code version, the above written masterSerializationVersion
		 * is first checked. Naively, the old code would throw an exception if a newer (i.e. higher) version is read (because
		 * the newer code could define a different byte stream even for the beginning of the file). However, newer code
		 * can write the earliest previous serialization version that is already capable to read this newer byte stream,
		 * despite not knowing about the additional content (i.e. typically since the changes in the versions since then
		 * are only appended at the end of the byte stream).
		 * 
		 * History:
		 * version 1, (selective upward compatibility not yet implemented)
		 * version 2, earliestSerializationVersionCapableOfReadingThis 2 (because sheets are added early in the byte stream)
		 */
		output.writeInt(masterSerializationVersion); // WARNING: if unsure, set this one to masterSerializationVersion
		
		
		// iterate all sheets
		output.writeInt(sheetStates.size());
		for (String sheetName : sheetStates.keySet()) {
			SheetState sheetState = sheetStates.get(sheetName);
			SerializationHelpers.writeNullableString(sheetName, output, masterSerializationVersion);
			
			// write cells map:
			output.writeInt(sheetState.cells.size());
			for (CellAddress address : sheetState.cells.keySet()) {
				SerializationHelpers.writeCellAddress(address, output, masterSerializationVersion);
				sheetState.cells.get(address).writeExternal(output, masterSerializationVersion);
			}
			
			output.writeBoolean(sheetState.freezeSheetAtTopLeftCornerOfCell != null);
			if (sheetState.freezeSheetAtTopLeftCornerOfCell != null)
				SerializationHelpers.writeCellAddress(sheetState.freezeSheetAtTopLeftCornerOfCell, output, masterSerializationVersion);
	
			output.writeBoolean(sheetState.autoFilterRange != null);
			if (sheetState.autoFilterRange != null)
				SerializationHelpers.writeCellRangeAddress(sheetState.autoFilterRange, output, masterSerializationVersion);
			
			output.writeInt(sheetState.rowHeights.size());
			for (Integer key : sheetState.rowHeights.keySet()) {
				output.writeInt(key);
				output.writeDouble(sheetState.rowHeights.get(key));
			}
			
			output.writeInt(sheetState.columnWidths.size());
			for (Integer key : sheetState.columnWidths.keySet()) {
				output.writeInt(key);
				Double columnWidth = sheetState.columnWidths.get(key);
				output.writeBoolean(columnWidth != null); // null means auto-size
				if (columnWidth != null)
					output.writeDouble(sheetState.columnWidths.get(key));
			}
			
			output.writeInt(sheetState.hiddenRows.size());
			for (Integer key : sheetState.hiddenRows)
				output.writeInt(key);
			
			output.writeInt(sheetState.hiddenColumns.size());
			for (Integer key : sheetState.hiddenColumns)
				output.writeInt(key);
			
			output.writeInt(sheetState.mergeRanges.size());
			for (CellRangeAddress mergeRange : sheetState.mergeRanges)
				SerializationHelpers.writeCellRangeAddress(mergeRange, output, masterSerializationVersion);
			
			output.writeInt(sheetState.columnGroups.size());
			for (Map.Entry<Pair<Integer, Integer>, Boolean> group : sheetState.columnGroups.entrySet())
				SerializationHelpers.writeGroupingEntry(group.getKey().getLeft(), group.getKey().getRight(), group.getValue(), output, masterSerializationVersion);
			
			output.writeInt(sheetState.rowGroups.size());
			for (Map.Entry<Pair<Integer, Integer>, Boolean> group : sheetState.rowGroups.entrySet())
				SerializationHelpers.writeGroupingEntry(group.getKey().getLeft(), group.getKey().getRight(), group.getValue(), output, masterSerializationVersion);
		}
	}

	@Override
	public PortObjectSpec getSpec() {
		return new XlsFormatterStateSpec();
	}

	@Override
	public JComponent[] getViews() {
		JPanel panel = new JPanel();

		panel.setName("XLS Formatter Port");
		panel.setLayout(new BorderLayout());

		try {
			JLabel label = new JLabel();
			label.setText(toBasicHtml("XLS Formatter Information (for debugging purposes only)", toLongString(true)));
			
			JScrollPane scrollPane = new JScrollPane(label, JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED, JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
			
			panel.add(scrollPane);
		} catch (Exception e) {
			e.printStackTrace();
		}

		return new JComponent[] { panel };
	}
	
	/**
	 * Returns a deep clone of the state via in-memory serialization and de-serialization.
	 */
	public XlsFormatterState getDeepClone() throws Exception {
		try {
			ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
			(new ObjectOutputStream(byteOutputStream)).writeObject(this);
	    return (XlsFormatterState)(new ObjectInputStream(new ByteArrayInputStream(byteOutputStream.toByteArray()))).readObject();
		} catch (Exception e) {
			throw new Exception("Implementation issue: Error during cloning of XLS Formatter State.", e);
		}
	}
	
	@Override
	public boolean equals(Object obj) {
		if (obj == null)
			return false;
		
		try {
			ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
			(new ObjectOutputStream(byteOutputStream)).writeObject(this);
			byte[] thisAsByteArray = byteOutputStream.toByteArray();
			
			byteOutputStream = new ByteArrayOutputStream();
			(new ObjectOutputStream(byteOutputStream)).writeObject(obj);
			byte[] refAsByteArray = byteOutputStream.toByteArray();
			
			if (thisAsByteArray.length != refAsByteArray.length)
				return false;
			
			for (int i = 0; i < thisAsByteArray.length; i++)
				if (thisAsByteArray[i] != refAsByteArray[i])
					return false;
			return true;
		} catch (IOException ioe) {
			return false;
		}
	}
	
	@Override
	public int hashCode() {
		try {
			ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
			(new ObjectOutputStream(byteOutputStream)).writeObject(this); 
			return byteOutputStream.toByteArray().hashCode();
		} catch (IOException ioe) {
			return 0;
		}
	}
	
	/**
	 * Returns a deep clone of the state passed as PortObject via in-memory serialization and de-serialization
	 * or a new XlsFormatterState, if portObject is null.
	 */
	public static XlsFormatterState getDeepClone(PortObject portObject) throws Exception {
		if (portObject == null)
			return new XlsFormatterState();
		return ((XlsFormatterState)portObject).getDeepClone();
	}
	
	
	/////////////////
	// Aux Functions
	
	/**
	 * Replace String UNMODIFIED in upper case and lower case writing by a slash (for more leaner visual output in the port view).
	 */
	private static String resolveEnumUnmodified(String value) {
		return value.replace("UNMODIFIED", "/").replace("unmodified", "/");
	}
	
	/**
	 * Convert a text to a basic HTML page for prettier output in port view).
	 */
	private static String toBasicHtml(String title, String text) {
		return "<html><body><b>" + title + "</b><p>" + text.replace("\n", "<br />") + "</p></html></body>"; 
	}
}
