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

package com.continental.knime.xlsformatter.xlscontroltablegenerator;

import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;

public class XlsControlTableGeneratorNodeDialog extends DefaultNodeSettingsPane {

	protected XlsControlTableGeneratorNodeDialog() {
		super();

		this.createNewGroup("Shift Rows Option");
		this.addDialogComponent(new DialogComponentBoolean(
				new SettingsModelBoolean(XlsControlTableGeneratorNodeModel.CFGKEY_ROWSHIFT, XlsControlTableGeneratorNodeModel.DEFAULT_ROWSHIFT),
				"write column header to first row"));

		this.createNewGroup("Result Table Structure Options");
		SettingsModelBoolean unpivot = new SettingsModelBoolean(XlsControlTableGeneratorNodeModel.CFGKEY_UNPIVOT, XlsControlTableGeneratorNodeModel.DEFAULT_UNPIVOT);
		this.addDialogComponent(new DialogComponentBoolean(
				unpivot, "unpivot result table (for easier post-processing and re-pivoting)"));

		SettingsModelBoolean extendedunpivotcolumns = new SettingsModelBoolean(XlsControlTableGeneratorNodeModel.CFGKEY_EXTENDEDUNPIVOTCOLUMNS, XlsControlTableGeneratorNodeModel.DEFAULT_EXTENDEDUNPIVOTCOLUMNS);
		this.addDialogComponent(new DialogComponentBoolean(
				extendedunpivotcolumns, "add more header columns"));
		
		unpivot.addChangeListener(new ChangeListener() {
			public void stateChanged(final ChangeEvent e) {
				extendedunpivotcolumns.setEnabled(unpivot.getBooleanValue());
			}
		});
	}
}

