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

package com.continental.knime.xlsformatter.xlscontroltablemerger;

import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentButtonGroup;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

import com.continental.knime.xlsformatter.commons.Commons.Modes;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;

public class XlsControlTableMergerNodeDialog extends DefaultNodeSettingsPane {

	protected XlsControlTableMergerNodeDialog() {
		super();

		String[] modeButtons = XlsFormatterUiOptions.getDropdownArrayFromEnum(Modes.values());
		addDialogComponent(new DialogComponentButtonGroup(
				new SettingsModelString(XlsControlTableMergerNodeModel.CFGKEY_MODE_STRING, XlsControlTableMergerNodeModel.DEFAULT_MODE_STRING), 
				"how to combine second with first table", false, modeButtons, modeButtons));
	}
}
