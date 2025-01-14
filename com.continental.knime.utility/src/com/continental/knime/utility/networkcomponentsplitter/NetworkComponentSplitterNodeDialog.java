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

package com.continental.knime.utility.networkcomponentsplitter;

import org.knime.core.data.DataColumnSpec;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentColumnNameSelection;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelColumnName;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.util.ColumnFilter;

public class NetworkComponentSplitterNodeDialog extends DefaultNodeSettingsPane {

	SettingsModelColumnName inputColumn1;
	SettingsModelColumnName inputColumn2;
	SettingsModelString outputColumnNameNode;
	SettingsModelString outputColumnNameCluster;
	
	protected NetworkComponentSplitterNodeDialog() {
		super();
		
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

		//create Node Selection Pane
		this.createNewGroup("Node Selection");
		inputColumn1 = new SettingsModelColumnName(NetworkComponentSplitterNodeModel.CFGKEY_COLUM_NAME1, NetworkComponentSplitterNodeModel.DEFAULT_COLUMN_NAME1);
		this.addDialogComponent(new DialogComponentColumnNameSelection( // Select Node1
				inputColumn1, "Select Node1 Column", 0, true, pureStringColumnFilter));
		inputColumn2 = new SettingsModelColumnName(NetworkComponentSplitterNodeModel.CFGKEY_COLUMN_NAME2,NetworkComponentSplitterNodeModel.DEFAULT_COLUMN_NAME2);
		this.addDialogComponent(new DialogComponentColumnNameSelection( // Select Node2
				inputColumn2, "Select Node2 Column", 0, true, pureStringColumnFilter));

		// create Missing Value Handling Pane
		this.createNewGroup("Missing Value Handling");
		this.addDialogComponent(new DialogComponentBoolean( // Missing Value Tickbox
				new SettingsModelBoolean(NetworkComponentSplitterNodeModel.CFGKEY_MISSING, NetworkComponentSplitterNodeModel.DEFAULT_MISSING),
				"Handle Missing Value as Node"));

		//create Output column Rename Pane
		this.createNewGroup("Output Column Names");
		outputColumnNameNode = new SettingsModelString(NetworkComponentSplitterNodeModel.CFGKEY_OUTPUT_COLUMN_NAME_NODE, NetworkComponentSplitterNodeModel.DEFAULT_OUTPUT_COLUMN_NAME_NODE);
		this.addDialogComponent(new DialogComponentString(outputColumnNameNode, "Nodelist Column Name"));        
		outputColumnNameCluster = new SettingsModelString(NetworkComponentSplitterNodeModel.CFGKEY_OUTPUT_COLUMN_NAME_CLUSTER, NetworkComponentSplitterNodeModel.DEFAULT_OUTPUT_COLUMN_NAME_CLUSTER);
		this.addDialogComponent(new DialogComponentString(outputColumnNameCluster, "ClusterID Column Name"));
	}
	
	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);
		
		if (inputColumn1.useRowID() || inputColumn2.useRowID())
			throw new InvalidSettingsException("RowID is not a valid column choice.");
		
		if (inputColumn1.getStringValue().equals(inputColumn2.getStringValue()))
			throw new InvalidSettingsException("Two different input columns need to be selected.");
		
		if (outputColumnNameNode.getStringValue().trim().equals(""))
			throw new InvalidSettingsException("The name of the first output column cannot be empty.");
		
		if (outputColumnNameCluster.getStringValue().trim().equals(""))
			throw new InvalidSettingsException("The name of the second output column cannot be empty.");
		
		if (outputColumnNameNode.getStringValue().trim().equals(outputColumnNameCluster.getStringValue().trim()))
			throw new InvalidSettingsException("Two different output columns names need to be defined.");
	}
}
