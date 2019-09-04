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

package com.continental.knime.xlsformatter.sheetselector;

import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

public class XlsSheetSelectorNodeDialog extends DefaultNodeSettingsPane {

	SettingsModelString sheetName;
	SettingsModelBoolean optionNamed;
	
	ChangeListener changeListener = new SheetSelectorDialogChangeListener();
	
	protected XlsSheetSelectorNodeDialog() {
		super();

		optionNamed = new SettingsModelBoolean(XlsSheetSelectorNodeModel.CFGKEY_OPTION_NAMED, XlsSheetSelectorNodeModel.DEFAULT_OPTION_NAMED);
		DialogComponentBoolean optionNamedComponent = new DialogComponentBoolean(optionNamed, "select sheet by name?");
		optionNamedComponent.setToolTipText("If deactivated, the first sheet is used instead (as if this node would be omitted).");
		this.addDialogComponent(optionNamedComponent);
		
		sheetName = new SettingsModelString(XlsSheetSelectorNodeModel.CFGKEY_SHEET_NAME, XlsSheetSelectorNodeModel.DEFAULT_SHEET_NAME);
		DialogComponentString sheetNameComponent = new DialogComponentString(sheetName, "sheet name", true, 20); 
		sheetNameComponent.setToolTipText("The name of the XLS sheet that you want further instructions to be applied on.");
		this.addDialogComponent(sheetNameComponent);
		
		optionNamed.addChangeListener(changeListener);
	}
	
	class SheetSelectorDialogChangeListener implements ChangeListener {
		public void stateChanged(final ChangeEvent e) {
			sheetName.setEnabled(optionNamed.getBooleanValue());
		}
	}
	
	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);
		if (optionNamed.getBooleanValue() && sheetName.getStringValue().trim().equals(""))
			throw new InvalidSettingsException("An empty sheet name is not allowed.");
	}
}
