package com.continental.knime.xlsformatter.porttype;

import java.util.Map;

import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeLogger;

import com.continental.knime.xlsformatter.commons.AddressingTools;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.BorderEdge;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellAlignmentHorizontal;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellAlignmentVertical;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellDataType;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.FillPattern;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.FormattingFlag;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.SheetState;

public class XlsFormatterStateMerger {
	
	/**
	 * Merges two XlsFormatterStates. In case of conflict, the master state is superior to the slave state.
	 * Will deeply clone the slave state first in order to safely re-use its contained objects in the master state.
	 * Slave can be null, in this case no action is performed. An exception is thrown if master is null.
	 */
	public static void mergeFormatterStates(XlsFormatterState master, XlsFormatterState slave, final ExecutionContext exec, final NodeLogger logger) throws Exception {
		
		if (master == null)
			throw new Exception("Coding issue: the master state cannot be null in a merge operation.");
			
		if (slave == null)
			return;
		
		slave = XlsFormatterState.getDeepClone(slave); // create fresh nested objects to safely add them to master later on
		
		// iterate over all sheets
		for (String slaveSheetName : slave.sheetStates.keySet()) {
			
			if (!master.sheetStates.containsKey(slaveSheetName))
				master.sheetStates.put(slaveSheetName, slave.sheetStates.get(slaveSheetName));
			else { // deeper merge needed as sheet is contained in both master and slave
				
				SheetState slaveState = slave.sheetStates.get(slaveSheetName);
				SheetState masterState = master.sheetStates.get(slaveSheetName);
				
				// merge the contained cells:
				for (CellAddress slaveCell : slaveState.cells.keySet())
					if (masterState.cells.containsKey(slaveCell))
						mergeCells(masterState.cells.get(slaveCell), slaveState.cells.get(slaveCell));
					else
						masterState.cells.put(slaveCell, slaveState.cells.get(slaveCell));
				
				// merge the sheet level formatting instructions:
				if (masterState.freezeSheetAtTopLeftCornerOfCell == null)
					masterState.freezeSheetAtTopLeftCornerOfCell = slaveState.freezeSheetAtTopLeftCornerOfCell;
				
				if (masterState.autoFilterRange == null)
					masterState.autoFilterRange = slaveState.autoFilterRange;
				
				for (Map.Entry<Integer, Double> entry : slaveState.columnWidths.entrySet())
					if (!masterState.columnWidths.containsKey(entry.getKey()))
						masterState.columnWidths.put(entry.getKey(), entry.getValue());
				
				for (Map.Entry<Integer, Double> entry : slaveState.rowHeights.entrySet())
					if (!masterState.rowHeights.containsKey(entry.getKey()))
						masterState.rowHeights.put(entry.getKey(), entry.getValue());
				
				masterState.hiddenRows.addAll(slaveState.hiddenRows);
				masterState.hiddenColumns.addAll(slaveState.hiddenColumns);
				
				// handle cell merge ranges by first eliminating pure duplicates and then checking for overlaps
				for (CellRangeAddress masterRange : masterState.mergeRanges)
					while (slaveState.mergeRanges.contains(masterRange))
						slaveState.mergeRanges.remove(masterRange);
				if (AddressingTools.hasOverlap(masterState.mergeRanges, slaveState.mergeRanges, exec, logger))
					throw new IllegalArgumentException("The provided XLS Formatter ports address sheet " +
							(slaveSheetName == null ? "[default / first]" : "\"" + slaveSheetName + "\"") +
							" which contain overlapping / conflicting cell merge ranges. See log for details.");
				masterState.mergeRanges.addAll(slaveState.mergeRanges);
				
				mergeGroups(masterState.columnGroups, slaveState.columnGroups);
				mergeGroups(masterState.rowGroups, slaveState.rowGroups);
			}
		}
	}
	
	/**
	 * Merges to CellStates. In case of conflict, master wins over slave, but at the lowest level of detail.
	 * Expects slave to be a fresh deep clone, i.e. its contained objects to be re-usable in master.
	 */
	private static void mergeCells(CellState master, CellState slave) {
		
		if (master.fontSize == null)
			master.fontSize = slave.fontSize;
		if (master.fontBold == FormattingFlag.UNMODIFIED)
			master.fontBold = slave.fontBold;
		if (master.fontItalic == FormattingFlag.UNMODIFIED)
			master.fontItalic = slave.fontItalic;
		if (master.fontUnderline == FormattingFlag.UNMODIFIED)
			master.fontUnderline = slave.fontUnderline;
		if (master.fontColor == null)
			master.fontColor = slave.fontColor;

		if (master.cellHorizontalAlignment == CellAlignmentHorizontal.UNMODIFIED)
			master.cellHorizontalAlignment = slave.cellHorizontalAlignment;
		if (master.cellVerticalAlignment == CellAlignmentVertical.UNMODIFIED)
			master.cellVerticalAlignment = slave.cellVerticalAlignment;
		if (master.wrapText == FormattingFlag.UNMODIFIED)
			master.wrapText = slave.wrapText;
		if (master.textTiltDegree == null)
			master.textTiltDegree = slave.textTiltDegree;
		
		if (master.backgroundColor == null)
			master.backgroundColor = slave.backgroundColor;
		if (master.fillPattern == FillPattern.UNMODIFIED)
			master.fillPattern = slave.fillPattern;
		if (master.fillForegroundColor == null)
			master.fillForegroundColor = slave.fillForegroundColor;
		
		/*
		 * Note that each cell has a pointer to a ConditionalFormattingSet. Within a range, these point to the same object.
		 * But even within a sheet, they might repeat content-wise. That's why only in the apply logic, they are checked
		 * for duplicates via toString() to generate just one POI sheetConditionalCf. Hence it is safe here to pass
		 * these objects through without any grouping.
		 */
		if (master.conditionalFormat == null)
			master.conditionalFormat = slave.conditionalFormat;
		
		if (master.cellDataType == CellDataType.UNMODIFIED)
			
			master.cellDataType = slave.cellDataType;
		if (master.textFormat == null)
			master.textFormat = slave.textFormat;
		if (master.hyperlink == null)
			master.hyperlink = slave.hyperlink;
		
		
		if (master.borderTop == null)
			master.borderTop = slave.borderTop;
		else
			BorderEdge.merge(master.borderTop, slave.borderTop);
		
		if (master.borderBottom == null)
			master.borderBottom = slave.borderBottom;
		else
			BorderEdge.merge(master.borderBottom, slave.borderBottom);
		
		if (master.borderLeft == null)
			master.borderLeft = slave.borderLeft;
		else
			BorderEdge.merge(master.borderLeft, slave.borderLeft);
		
		if (master.borderRight == null)
			master.borderRight = slave.borderRight;
		else
			BorderEdge.merge(master.borderRight, slave.borderRight);
		
		if (master.borderDiagonalSlash == null)
			master.borderDiagonalSlash = slave.borderDiagonalSlash;
		else
			BorderEdge.merge(master.borderDiagonalSlash, slave.borderDiagonalSlash);
		
		if (master.borderDiagonalBackslash == null)
			master.borderDiagonalBackslash = slave.borderDiagonalBackslash;
		else
			BorderEdge.merge(master.borderDiagonalBackslash, slave.borderDiagonalBackslash);
	}
	
	/**
	 * Merges the data structure used for column and row groups. The maps consist of <<from,to>,isCollapsed>.
	 * As overlapping group ranges are allowed, only exactly overlapping ranges differing in the collapsed
	 * flag are overwritten. All other cases are added from slave to master.
	 */
	private static void mergeGroups(
			Map<Pair<Integer, Integer>, Boolean> master,
			Map<Pair<Integer, Integer>, Boolean> slave) throws Exception {
		
		for (Map.Entry<Pair<Integer, Integer>, Boolean> slaveEntry : slave.entrySet())
			if (!master.containsKey(slaveEntry.getKey()))
				master.put(slaveEntry.getKey(), slaveEntry.getValue());
	}
}
