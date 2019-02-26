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

package com.continental.knime.xlsformatter.borderformatter;

import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.util.CellAddress;
import org.knime.core.node.ExecutionContext;

import com.continental.knime.xlsformatter.commons.AddressingTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.BorderEdge;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.BorderStyle;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellState;

public class XlsFormatterBorderFormatterLogic {

	private enum BorderPosition { TOP, LEFT, BOTTOM, RIGHT, INNER_VERTICAL, INNER_HORIZONTAL }
	
	/**
	 * The cell-based part of a XlsFormatterState mapping cell addresses to cell states.
	 * Will be modified by the border logic.
	 */
	private Map<CellAddress, CellState> _cellsMap;
	
	/**
	 * A list of cell addresses that match the searched tag, hence representing the border target area.
	 */
	private List<CellAddress> _matchingCells;
	
	/**
	 * The requested border format (style and color).
	 */
	private XlsFormatterState.BorderEdge _borderEdge;
	
	/**
	 * Contains same content as matchingCells list, but for optimized look-up speed as hashset.
	 */
	private Set<CellAddress> _matchingCellsSet;
	
	/**
	 * 
	 * @param cellsMap The cell-based part of a XlsFormatterState mapping cell addresses to cell states.
	 * @param matchingCells A list of cell addresses that match the searched tag, hence representing the border target area.
	 * @param borderEdge The requested border format (style and color).
	 */
	public XlsFormatterBorderFormatterLogic(Map<CellAddress, CellState> cellsMap, List<CellAddress> matchingCells,
			XlsFormatterState.BorderEdge borderEdge) {
		_cellsMap = cellsMap;
		_matchingCells = matchingCells;
		_borderEdge = borderEdge;
		_matchingCellsSet = new HashSet<CellAddress>(matchingCells);
	}
	
	// prevent creation of this class without using the above constructor with arguments
	@SuppressWarnings("unused")
	private XlsFormatterBorderFormatterLogic() { }
	
	/**
	 * Update the linked cell address to state map to consider all borders derived from the linked matching cells list.
	 */
	public void implementBordersMatchingTagsInMap(
			final boolean top, final boolean left, final boolean bottom, final boolean right,
			final boolean innerVertical, final boolean innerHorizontal, final ExecutionContext exec) throws Exception {

		if (_borderEdge == null || (_borderEdge.style == BorderStyle.UNMODIFIED && _borderEdge.color == null)) // nothing to do
				return;
		
		for (CellAddress cell : _matchingCells) {
			exec.checkCanceled();
			
			if (top)
				setCellBorderAndNeighbors(cell, BorderPosition.TOP);
			if (left)
				setCellBorderAndNeighbors(cell, BorderPosition.LEFT);
			if (bottom)
				setCellBorderAndNeighbors(cell, BorderPosition.BOTTOM);
			if (right)
				setCellBorderAndNeighbors(cell, BorderPosition.RIGHT);
			if (innerVertical)
				setCellBorderAndNeighbors(cell, BorderPosition.INNER_VERTICAL);
			if (innerHorizontal)
				setCellBorderAndNeighbors(cell, BorderPosition.INNER_HORIZONTAL);
		}
	}
	
	/**
	 * Set a cell border at a time (incl. its corresponding neighbor cell)
	 */
	private void setCellBorderAndNeighbors(final CellAddress cell, final BorderPosition borderPosition) throws Exception {

		switch (borderPosition) {
		case TOP: // if there is no further matching cell above, set this outer top border (and the respective identical one of the top neighbor cell)
			if (!isRelativeCellAMatch(cell, 0, -1)) {
				updateStatesBorder(AddressingTools.safelyGetCellInMap(_cellsMap, cell), BorderPosition.TOP);
				setRelativePositionedSharedCellBorderEdge(cell, 0, -1);
			}
			break;

		case BOTTOM: // if there is no further matching cell underneath, set this outer bottom border (and the respective identical one of the bottom neighbor cell)
			if (!isRelativeCellAMatch(cell, 0, 1)) {
				updateStatesBorder(AddressingTools.safelyGetCellInMap(_cellsMap, cell), BorderPosition.BOTTOM);
				setRelativePositionedSharedCellBorderEdge(cell, 0, 1);
			}
			break;

		case LEFT: // if there is no further matching cell to the left, set this outer left border (and the respective identical one of the left neighbor cell)
			if (!isRelativeCellAMatch(cell, -1, 0)) {
				updateStatesBorder(AddressingTools.safelyGetCellInMap(_cellsMap, cell), BorderPosition.LEFT);
				setRelativePositionedSharedCellBorderEdge(cell, -1, 0);
			}
			break;

		case RIGHT: // if there is no further matching cell to the right, set this outer right border (and the respective identical one of the right neighbor cell)
			if (!isRelativeCellAMatch(cell, 1, 0)) {
				updateStatesBorder(AddressingTools.safelyGetCellInMap(_cellsMap, cell), BorderPosition.RIGHT);
				setRelativePositionedSharedCellBorderEdge(cell, 1, 0);
			}
			break;

		case INNER_VERTICAL: // if there is an also matching neighbor cell to the left/right, set this inner left/right border (and the respective identical one of the left/right neighbor cell)
			if (isRelativeCellAMatch(cell, -1, 0)) {
				updateStatesBorder(AddressingTools.safelyGetCellInMap(_cellsMap, cell), BorderPosition.LEFT);
				setRelativePositionedSharedCellBorderEdge(cell, -1, 0);
			}
			if (isRelativeCellAMatch(cell, 1, 0)) {
				updateStatesBorder(AddressingTools.safelyGetCellInMap(_cellsMap, cell), BorderPosition.RIGHT);
				setRelativePositionedSharedCellBorderEdge(cell, 1, 0);
			}
			break;
			
		case INNER_HORIZONTAL: // if there is an also matching neighbor cell above/underneath, set this inner top/bottom border (and the respective identical one of the top/bottom neighbor cell)
			if (isRelativeCellAMatch(cell, 0, -1)) {
				updateStatesBorder(AddressingTools.safelyGetCellInMap(_cellsMap, cell), BorderPosition.TOP);
				setRelativePositionedSharedCellBorderEdge(cell, 0, -1);
			}
			if (isRelativeCellAMatch(cell, 0, 1)) {
				updateStatesBorder(AddressingTools.safelyGetCellInMap(_cellsMap, cell), BorderPosition.BOTTOM);
				setRelativePositionedSharedCellBorderEdge(cell, 0, 1);
			}
			break;
			
		default:
			throw new Exception("Coding issue: unknown border position.");
		}
	}
	
	
	/**
	 * Logic to set every border definition for both cells sharing this edge.
	 * (E.g. cell B2 top border deliberately formatted means B1 bottom border set via this method as well.)
	 */
	private void setRelativePositionedSharedCellBorderEdge(final CellAddress cell,
			final int diffX, final int diffY) throws Exception {
		
		CellState neighborState = AddressingTools.safelyGetCellWithNegativeCheck(_cellsMap, cell.getColumn() + diffX, cell.getRow() + diffY);
		if (neighborState != null) { // null would mean e.g. we looked left of cell A1
			if (diffX == -1 && diffY == 0)
				updateStatesBorder(neighborState, BorderPosition.RIGHT);
			else if (diffX == 1 && diffY == 0)
				updateStatesBorder(neighborState, BorderPosition.LEFT);
			else if (diffX == 0 && diffY == -1)
				updateStatesBorder(neighborState, BorderPosition.BOTTOM);
			else if (diffX == 0 && diffY == 1)
				updateStatesBorder(neighborState, BorderPosition.TOP);
			else
				throw new Exception("Coding issue: relative border setting with invalid jump.");
		}
	}
	
	
	
	private boolean isRelativeCellAMatch(final CellAddress cell, final int diffX, final int diffY) {
		if (!XlsFormatterControlTableValidator.isCellWithinXlsSpec(cell.getRow() + diffY, cell.getColumn() + diffX)) // e.g. left neighbor of cell A1 -> cannot be a match
			return false;
		return _matchingCellsSet.contains(new CellAddress(cell.getRow() + diffY, cell.getColumn() + diffX));
	}
	
	
	private void updateStatesBorder(CellState state, BorderPosition position) throws Exception {
		switch (position) {
		case TOP:
			if (state.borderTop == null)
				state.borderTop = new BorderEdge(_borderEdge.style, _borderEdge.color);
			else
				state.borderTop.mergeIn(_borderEdge);
			break;
		case BOTTOM:
			if (state.borderBottom == null)
				state.borderBottom = new BorderEdge(_borderEdge.style, _borderEdge.color);
			else
				state.borderBottom.mergeIn(_borderEdge);
			break;
		case LEFT:
			if (state.borderLeft == null)
				state.borderLeft = new BorderEdge(_borderEdge.style, _borderEdge.color);
			else
				state.borderLeft.mergeIn(_borderEdge);
			break;
		case RIGHT:
			if (state.borderRight == null)
				state.borderRight = new BorderEdge(_borderEdge.style, _borderEdge.color);
			else
				state.borderRight.mergeIn(_borderEdge);
			break;
		default:
			throw new Exception("Coding issue: Inner border positions are not stored directly in state.");
		}
	}
}
