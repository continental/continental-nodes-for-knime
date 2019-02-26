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

package com.continental.knime.xlsformatter.xlscontroltablefromcellrange;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentButtonGroup;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

import com.continental.knime.xlsformatter.commons.AddressingTools;
import com.continental.knime.xlsformatter.commons.Commons.Modes;
import com.continental.knime.xlsformatter.commons.UiValidation;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;

public class XlsControlTableFromCellRangeNodeDialog extends DefaultNodeSettingsPane {

	SettingsModelString cellRange;
	SettingsModelString tag;
	SettingsModelString combineMode;

	protected XlsControlTableFromCellRangeNodeDialog() {
		super();

		cellRange = new SettingsModelString(XlsControlTableFromCellRangeNodeModel.CFGKEY_CELLRANGE_STRING, XlsControlTableFromCellRangeNodeModel.DEFAULT_CELLRANGE_STRING);
		addDialogComponent(new DialogComponentString(cellRange, "cell range"));

		tag = new SettingsModelString(XlsControlTableFromCellRangeNodeModel.CFGKEY_TAG_STRING, XlsControlTableFromCellRangeNodeModel.DEFAULT_TAG_STRING);
		addDialogComponent(new DialogComponentString(tag, "tag to set in control table"));

		String[] modeButtons = XlsFormatterUiOptions.getDropdownArrayFromEnum(Modes.values());
		combineMode = new SettingsModelString(XlsControlTableFromCellRangeNodeModel.CFGKEY_MODE_STRING, XlsControlTableFromCellRangeNodeModel.DEFAULT_MODE_STRING);
		addDialogComponent(new DialogComponentButtonGroup(combineMode, "combine mode", false, modeButtons, modeButtons));
	}

	
	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);

		UiValidation.validateTagField(tag);

		try {
			AddressingTools.parseRange(cellRange.getStringValue().trim());
		} catch (IllegalArgumentException iae) {
			throw new InvalidSettingsException(iae.getMessage(), iae);
		}
	}
}

