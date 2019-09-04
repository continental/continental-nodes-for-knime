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

package com.continental.knime.xlsformatter.sheetproperties;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentButtonGroup;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

import com.continental.knime.xlsformatter.commons.UiValidation;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;

public class XlsSheetPropertiesNodeDialog extends DefaultNodeSettingsPane {

	private static String[] functionOptions = XlsFormatterUiOptions.getDropdownArrayFromEnum(XlsSheetPropertiesNodeModel.FunctionOptions.values());

	SettingsModelString tag;
	
	protected XlsSheetPropertiesNodeDialog() {
		super();

		this.createNewGroup("Tag Selection");
		tag = new SettingsModelString(XlsSheetPropertiesNodeModel.CFGKEY_TAG, XlsSheetPropertiesNodeModel.DEFAULT_TAG);
		DialogComponentString tagComponent = new DialogComponentString(tag, XlsFormatterUiOptions.UI_LABEL_SINGLE_TAG, true, 10); 
		tagComponent.setToolTipText(XlsFormatterUiOptions.UI_TOOLTIP_SINGLE_TAG);
		this.addDialogComponent(tagComponent);

		
		this.createNewGroup("Sheet Propoerties");
		SettingsModelString function = new SettingsModelString(XlsSheetPropertiesNodeModel.CFGKEY_FUNCTION, XlsSheetPropertiesNodeModel.DEFAULT_FUNCTION);
		DialogComponentButtonGroup functionComponent = new DialogComponentButtonGroup( //ControlTableStyle
				function, "", true, functionOptions, functionOptions);
		functionComponent.setToolTipText("Select the function you want to apply in your XLS Sheet.");
		this.addDialogComponent(functionComponent);
	}
	
	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);
		UiValidation.validateTagField(tag);
	}
}
