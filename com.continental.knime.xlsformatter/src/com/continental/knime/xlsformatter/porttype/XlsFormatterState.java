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
import org.apache.commons.lang3.tuple.Triple;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.port.PortType;
import org.knime.core.node.port.PortTypeRegistry;

import com.continental.knime.xlsformatter.commons.ColorTools;


public class XlsFormatterState implements PortObject, Externalizable {

	/**
	 * A threshold value of how many cells are to be included in a String representation of a XlsFormatterState.
	 * Since performance on real-life cases with thousands of formatted cells would be too low, a cut-off
	 * value of 200 is recommended. 
	 */
	public static int VIEW_CELLS_THRESHOLD = 200;
	
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
		
		public static ConditionalFormattingSet readFromExternal(ObjectInput input, long masterSerialVersionUID) throws IOException, ClassNotFoundException {
			int tempSize = input.readInt();
			ConditionalFormattingSet ret = new ConditionalFormattingSet();
			ret.backgroundScaleFixpoints = new ArrayList<Pair<Double, Color>>(tempSize);
			for (int i = 0; i < tempSize; i++) {
				double step = input.readDouble();
				Color color = SerializationHelpers.readNullableColor(input, masterSerialVersionUID);
				ret.backgroundScaleFixpoints.add(Pair.of(step, color));
			}
			return ret;
		}

		public void writeExternal(ObjectOutput output, long masterSerialVersionUID) throws IOException {
			output.writeInt(backgroundScaleFixpoints.size());
			for (Pair<Double, Color> pair : backgroundScaleFixpoints) {
				output.writeDouble(pair.getLeft());
				SerializationHelpers.writeNullableColor(pair.getRight(), output, masterSerialVersionUID);
			}
		}
	}
	
	/**
	 * Formatting instructions of one edge of one cell's border
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
		
		public static BorderEdge readFromExternal(ObjectInput input, long masterSerialVersionUID) throws IOException, ClassNotFoundException {
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
			Color color = SerializationHelpers.readNullableColor(input, masterSerialVersionUID);
			return new BorderEdge(style, color);
		}

		public void writeExternal(ObjectOutput output, long masterSerialVersionUID) throws IOException {
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
			SerializationHelpers.writeNullableColor(color, output, masterSerialVersionUID);
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
		 */
		public String cellFormatToShortString(boolean restrictToPoiStyleRelevantProperties) {
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
			sb.append(";borders:T:");
			sb.append(borderTop == null ? "-;" : borderTop.toString());
			sb.append("B:");
			sb.append(borderBottom == null ? "-;" : borderBottom.toString());
			sb.append("L:");
			sb.append(borderLeft == null ? "-;" : borderLeft.toString());
			sb.append("R:");
			sb.append(borderRight == null ? "-;" : borderRight.toString());
			if (!restrictToPoiStyleRelevantProperties) {
				sb.append(conditionalFormat == null ? "cf:-" : conditionalFormat.toString());
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
				nonFormattingStateString = (new CellState()).cellFormatToShortString(true); 
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

		public static CellState readFromExternal(ObjectInput input, long masterSerialVersionUID) throws IOException, ClassNotFoundException {
			CellState ret = new CellState();
			ret.fontSize = SerializationHelpers.readNullableInt(input, masterSerialVersionUID);
			ret.fontBold = getFormattingFlagFromSerializedByte(input.readByte());
			ret.fontItalic = getFormattingFlagFromSerializedByte(input.readByte());
			ret.fontUnderline = getFormattingFlagFromSerializedByte(input.readByte());
			ret.fontColor = SerializationHelpers.readNullableColor(input, masterSerialVersionUID);
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
			ret.textTiltDegree = SerializationHelpers.readNullableInt(input, masterSerialVersionUID);
			ret.backgroundColor = SerializationHelpers.readNullableColor(input, masterSerialVersionUID);
			ret.fillForegroundColor = SerializationHelpers.readNullableColor(input, masterSerialVersionUID);
			ret.textFormat = SerializationHelpers.readNullableString(input, masterSerialVersionUID);
			ret.hyperlink = SerializationHelpers.readNullableString(input, masterSerialVersionUID);
			input.readByte(); // placeholder for future hyperlink type implementation, 0 meaning DEFAULT
			if (input.readBoolean())
				ret.conditionalFormat = ConditionalFormattingSet.readFromExternal(input, masterSerialVersionUID);
			if (input.readBoolean())
				ret.borderTop = BorderEdge.readFromExternal(input, masterSerialVersionUID);
			if (input.readBoolean())
				ret.borderBottom = BorderEdge.readFromExternal(input, masterSerialVersionUID);
			if (input.readBoolean())
				ret.borderLeft = BorderEdge.readFromExternal(input, masterSerialVersionUID);
			if (input.readBoolean())
				ret.borderRight = BorderEdge.readFromExternal(input, masterSerialVersionUID);
			if (input.readBoolean())
				ret.borderDiagonalSlash = BorderEdge.readFromExternal(input, masterSerialVersionUID);
			if (input.readBoolean())
				ret.borderDiagonalBackslash = BorderEdge.readFromExternal(input, masterSerialVersionUID);
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
		 * @param output
		 * @param masterSerialVersionUID
		 * @throws IOException
		 */
		public void writeExternal(ObjectOutput output, long masterSerialVersionUID) throws IOException {
			SerializationHelpers.writeNullableInt(fontSize, output, masterSerialVersionUID);
			output.writeByte(fontBold == FormattingFlag.UNMODIFIED ? 0 : (fontBold == FormattingFlag.OFF ? 1 : 2));
			output.writeByte(fontItalic == FormattingFlag.UNMODIFIED ? 0 : (fontItalic == FormattingFlag.OFF ? 1 : 2));
			output.writeByte(fontUnderline == FormattingFlag.UNMODIFIED ? 0 : (fontUnderline == FormattingFlag.OFF ? 1 : 2));
			SerializationHelpers.writeNullableColor(fontColor, output, masterSerialVersionUID);
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
			SerializationHelpers.writeNullableInt(textTiltDegree, output, masterSerialVersionUID);
			SerializationHelpers.writeNullableColor(backgroundColor, output, masterSerialVersionUID);
			SerializationHelpers.writeNullableColor(fillForegroundColor, output, masterSerialVersionUID);
			SerializationHelpers.writeNullableString(textFormat, output, masterSerialVersionUID);
			SerializationHelpers.writeNullableString(hyperlink, output, masterSerialVersionUID);
			output.writeByte(0); // placeholder for future hyperlink type implementation, 0 meaning DEFAULT
			
			output.writeBoolean(conditionalFormat != null);
			if (conditionalFormat != null)
				conditionalFormat.writeExternal(output, masterSerialVersionUID);
			
			for (BorderEdge border : new BorderEdge[] { borderTop, borderBottom, borderLeft, borderRight, borderDiagonalSlash, borderDiagonalBackslash }) {
				output.writeBoolean(border != null);
				if (border != null)
					border.writeExternal(output, masterSerialVersionUID);
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
	 * The serialization ID controlling backward compatibility for future releases.
	 * It is used as the "one and only" master serial version UID, even for subclasses.
	 */
	private static final long serialVersionUID = 1L;
	
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
	 * List of triples for column groups: fromCol, toCol, isCollapsed
	 */
	public List<Triple<Integer, Integer, Boolean>> columnGroups = new ArrayList<Triple<Integer, Integer, Boolean>>();
	
	/**
	 * List of triples for row groups: fromRow, toRow, isCollapsed
	 */
	public List<Triple<Integer, Integer, Boolean>> rowGroups = new ArrayList<Triple<Integer, Integer, Boolean>>();
	
	/**
	 * XlsFormatterState constructor without any action as all relevant fields are instantiated.
	 */
	public XlsFormatterState() { }
	
	/**
	 * Checks whether this state is empty, meaning all data instruction-storing structures are empty.
	 * @return
	 */
	public boolean isEmpty() {
		return cells.size() == 0 && freezeSheetAtTopLeftCornerOfCell == null && autoFilterRange == null &&
				rowHeights.size() == 0 && columnWidths.size() == 0 && hiddenRows.size() == 0 && hiddenColumns.size() == 0 &&
				mergeRanges.size() == 0 && columnGroups.size() == 0 && rowGroups.size() == 0;
	}
	
	/**
	 * Gets a full specification, e.g. for viewing the port state in the UI
	 */
	public String toLongString(boolean cutLongText) {
		StringBuilder sb = new StringBuilder();
		sb.append("\nfreezeSheetAtTopLeftCornerOfCell: ");
		sb.append(freezeSheetAtTopLeftCornerOfCell == null ? "-" : freezeSheetAtTopLeftCornerOfCell.toString());
		sb.append("\nautoFilterRange: ");
		sb.append(autoFilterRange == null ? "-" : autoFilterRange.formatAsString());
		sb.append("\nrowHeights: ");
		for (int i : rowHeights.keySet())
			sb.append(i + ":" +  rowHeights.get(i));
		sb.append("\ncolumnWidths: ");
		for (int i : columnWidths.keySet())
			sb.append(CellReference.convertNumToColString(i) + ":" + columnWidths.get(i) + " ");
		sb.append("\nhiddenRows: ");
		for (int i : hiddenRows)
			sb.append(i + " ");
		sb.append("\nhiddenColumns: ");
		for (int i : hiddenColumns)
			sb.append(i + " ");
		sb.append("\nmergeRanges: ");
		for (CellRangeAddress range : mergeRanges)
			sb.append(range.formatAsString() + " ");
		sb.append("\ncolumnGroups: ");
		for (Triple<Integer, Integer, Boolean> group : columnGroups)
			sb.append(CellReference.convertNumToColString(group.getLeft()) + ":" + CellReference.convertNumToColString(group.getMiddle()) + "," + (group.getRight() ? "collapsed" : "opened") + " ");
		sb.append("\nrowGroups: ");
		for (Triple<Integer, Integer, Boolean> group : columnGroups)
			sb.append(group.getLeft() + ":" + group.getMiddle() + "," + (group.getRight() ? "collapsed" : "opened") + " ");
		sb.append("\n");
		int iteration = 0;
		for (CellAddress address : cells.keySet()) {
			sb.append("\n" + address.formatAsString() + ": ");
			sb.append(cells.get(address).cellFormatToShortString(false));
			if (iteration++ >= VIEW_CELLS_THRESHOLD) {
				sb.append("\n[...], " + cells.size() + " total cells with instructions");
				break;
			}
		}
		
		return sb.toString();
	}
	


	@Override
	public void readExternal(ObjectInput input) throws IOException, ClassNotFoundException {
		int tempSize;
		int temp;
		long readSerialVersionUID = input.readLong();
		
		// read cells map:
		tempSize = input.readInt();
		cells = new HashMap<CellAddress, CellState>(tempSize);
		for (int i = 0; i < tempSize; i++) {
			CellAddress address = SerializationHelpers.readCellAddress(input, readSerialVersionUID);
			cells.put(
					address,
					CellState.readFromExternal(input, readSerialVersionUID));
		}
		
		if (input.readBoolean())
			freezeSheetAtTopLeftCornerOfCell = SerializationHelpers.readCellAddress(input, readSerialVersionUID);
		
		if (input.readBoolean())
			autoFilterRange = SerializationHelpers.readCellRangeAddress(input, readSerialVersionUID);
		
		tempSize = input.readInt();
		rowHeights = new HashMap<Integer, Double>(tempSize);
		for (int i = 0; i < tempSize; i++) {
			temp = input.readInt();
			rowHeights.put(temp, input.readDouble());
		}
		
		tempSize = input.readInt();
		columnWidths = new HashMap<Integer, Double>(tempSize);
		for (int i = 0; i < tempSize; i++) {
			temp = input.readInt(); // key, i.e. column index
			boolean autoSizeColumn = !input.readBoolean();
			Double columnWidth = null;
			if (!autoSizeColumn)
				columnWidth = input.readDouble();
			columnWidths.put(temp, columnWidth);
		}
		
		tempSize = input.readInt();
		hiddenRows = new HashSet<Integer>(tempSize);
		for (int i = 0; i < tempSize; i++)
			hiddenRows.add(input.readInt());
		
		tempSize = input.readInt();
		hiddenColumns = new HashSet<Integer>(tempSize);
		for (int i = 0; i < tempSize; i++)
			hiddenColumns.add(input.readInt());
		
		tempSize = input.readInt();
		mergeRanges = new ArrayList<CellRangeAddress>(tempSize);
		for (int i = 0; i < tempSize; i++)
			mergeRanges.add(SerializationHelpers.readCellRangeAddress(input, readSerialVersionUID));
		
		tempSize = input.readInt();
		columnGroups = new ArrayList<Triple<Integer, Integer, Boolean>>();
		for (int i = 0; i < tempSize; i++)
			columnGroups.add(SerializationHelpers.readGroupTriple(input, readSerialVersionUID));
		
		tempSize = input.readInt();
		rowGroups = new ArrayList<Triple<Integer, Integer, Boolean>>();
		for (int i = 0; i < tempSize; i++)
			rowGroups.add(SerializationHelpers.readGroupTriple(input, readSerialVersionUID));
	}

	
	
	@Override
	public void writeExternal(ObjectOutput output) throws IOException {
		
		output.writeLong(serialVersionUID);
		
		// write cells map:
		output.writeInt(cells.size());
		for (CellAddress address : cells.keySet()) {
			SerializationHelpers.writeCellAddress(address, output, serialVersionUID);
			cells.get(address).writeExternal(output, serialVersionUID);
		}
		
		output.writeBoolean(freezeSheetAtTopLeftCornerOfCell != null);
		if (freezeSheetAtTopLeftCornerOfCell != null)
			SerializationHelpers.writeCellAddress(freezeSheetAtTopLeftCornerOfCell, output, serialVersionUID);

		output.writeBoolean(autoFilterRange != null);
		if (autoFilterRange != null)
			SerializationHelpers.writeCellRangeAddress(autoFilterRange, output, serialVersionUID);
		
		output.writeInt(rowHeights.size());
		for (Integer key : rowHeights.keySet()) {
			output.writeInt(key);
			output.writeDouble(rowHeights.get(key));
		}
		
		output.writeInt(columnWidths.size());
		for (Integer key : columnWidths.keySet()) {
			output.writeInt(key);
			Double columnWidth = columnWidths.get(key);
			output.writeBoolean(columnWidth != null); // null means auto-size
			if (columnWidth != null)
				output.writeDouble(columnWidths.get(key));
		}
		
		output.writeInt(hiddenRows.size());
		for (Integer key : hiddenRows)
			output.writeInt(key);
		
		output.writeInt(hiddenColumns.size());
		for (Integer key : hiddenColumns)
			output.writeInt(key);
		
		output.writeInt(mergeRanges.size());
		for (CellRangeAddress mergeRange : mergeRanges)
			SerializationHelpers.writeCellRangeAddress(mergeRange, output, serialVersionUID);
		
		output.writeInt(columnGroups.size());
		for (Triple<Integer, Integer, Boolean> group : columnGroups)
			SerializationHelpers.writeGroupTriple(group, output, serialVersionUID);
		
		output.writeInt(rowGroups.size());
		for (Triple<Integer, Integer, Boolean> group : rowGroups)
			SerializationHelpers.writeGroupTriple(group, output, serialVersionUID);
	}
	
	@Override
	public String getSummary() {
		return "XLS Formatter Port";
	}

	@Override
	public PortObjectSpec getSpec() {
		XlsFormatterStateSpec spec = new XlsFormatterStateSpec();
		spec.setContainsMergeInstruction(mergeRanges.size() != 0);
		return spec;
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
