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

package com.continental.knime.utility.fiforesolver;

import org.knime.core.data.def.StringCell;
import org.knime.core.data.IntValue;
import org.knime.core.data.LongValue;
import org.knime.core.data.DataColumnSpec;
import org.knime.core.data.DoubleValue;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentButtonGroup;
import org.knime.core.node.defaultnodesettings.DialogComponentColumnNameSelection;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelColumnName;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.util.ColumnFilter;


public class FifoResolverNodeDialog extends DefaultNodeSettingsPane {

	SettingsModelString mode;
	SettingsModelColumnName inputColumnGroup;
	SettingsModelColumnName inputColumnQty;
	SettingsModelBoolean failAtInconsistency;
	
	@SuppressWarnings("unchecked")
	protected FifoResolverNodeDialog() {
		super();
		
		// create mode pane
		this.createNewGroup("Type of Queue to Resolve");
		setHorizontalPlacement(true);
		mode = new SettingsModelString(FifoResolverNodeModel.CFGKEY_MODE, FifoResolverNodeModel.DEFAULT_MODE);
		
		String[] modeButtonArray = { FifoResolverNodeModel.OPTION_MODE_FIFO, FifoResolverNodeModel.OPTION_MODE_LIFO };
		DialogComponentButtonGroup modeComponent = new DialogComponentButtonGroup(
				mode, "", false, modeButtonArray, modeButtonArray);
		modeComponent.setToolTipText("Choose one of the queue resolution modes 'first-in-first-out' (FIFO) or 'last-in-first-out' (LIFO).");
		this.addDialogComponent(modeComponent);
		setHorizontalPlacement(false);
		
		ColumnFilter pureStringColumnFilter = new ColumnFilter() {
            @Override
            public boolean includeColumn(DataColumnSpec columnSpec) {
                return columnSpec.getType().getCellClass().equals(StringCell.class);
            }
            @Override
            public String allFilteredMsg() {
                return "No String columns available.";
            }
        };
		
		// create Column Selection Pane
		this.createNewGroup("Column Selection");
		inputColumnGroup = new SettingsModelColumnName(FifoResolverNodeModel.CFGKEY_COLUM_NAME_GROUP, FifoResolverNodeModel.DEFAULT_COLUMN_NAME_GROUP);
		this.addDialogComponent(new DialogComponentColumnNameSelection(
				inputColumnGroup, "Select Grouping Column", 0, true, pureStringColumnFilter));
		inputColumnQty = new SettingsModelColumnName(FifoResolverNodeModel.CFGKEY_COLUMN_NAME_QTY, FifoResolverNodeModel.DEFAULT_COLUMN_NAME_QTY);
		this.addDialogComponent(new DialogComponentColumnNameSelection(
				inputColumnQty, "Select Quantity Column (positive values mean INCOMING, negative OUTGOING)", 0, true, IntValue.class, LongValue.class, DoubleValue.class));

		// create Missing Value Handling Pane
		this.createNewGroup("Exception Handling");
		failAtInconsistency = new SettingsModelBoolean(FifoResolverNodeModel.CFGKEY_FAIL_AT_INCONSISTENCY, FifoResolverNodeModel.DEFAULT_FAIL_AT_INCONSISTENCY);
		this.addDialogComponent(new DialogComponentBoolean(
				failAtInconsistency, "Fail at queuing inconsistency, i.e. more OUT than prior IN"));
	}
}
