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

package com.continental.knime.xlsformatter.commenter;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObjectSpec;

public class XlsFormatterCommenterNodeDialog extends DefaultNodeSettingsPane {

	SettingsModelString author;
		
	protected XlsFormatterCommenterNodeDialog() {
		super();

		author = new SettingsModelString(XlsFormatterCommenterNodeModel.CFGKEY_AUTHOR, XlsFormatterCommenterNodeModel.DEFAULT_AUTHOR);
		this.addDialogComponent(new DialogComponentString(author, "Author of the comment", false, 15));        

	}
	
	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);
	}
	
	@Override
	public void loadAdditionalSettingsFrom(NodeSettingsRO settings, PortObjectSpec[] specs)
			throws NotConfigurableException {
		super.loadAdditionalSettingsFrom(settings, specs);
	}
}
