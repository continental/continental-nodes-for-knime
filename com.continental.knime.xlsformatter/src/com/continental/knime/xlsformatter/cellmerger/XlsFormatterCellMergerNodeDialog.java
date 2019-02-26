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

package com.continental.knime.xlsformatter.cellmerger;

import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

import com.continental.knime.xlsformatter.commons.UiValidation;

public class XlsFormatterCellMergerNodeDialog extends DefaultNodeSettingsPane {

	SettingsModelString tag;
	SettingsModelBoolean alltags;

	protected XlsFormatterCellMergerNodeDialog() {
		super();

		this.createNewGroup("Tag Selection");
		this.setHorizontalPlacement(true);
		tag = new SettingsModelString(XlsFormatterCellMergerNodeModel.CFGKEY_TAGSTRING,XlsFormatterCellMergerNodeModel.DEFAULT_TAGSTRING);
		this.addDialogComponent(new DialogComponentString(tag, "applies to tag (single tag only)"));
		
		alltags = new SettingsModelBoolean(XlsFormatterCellMergerNodeModel.CFGKEY_ALL_TAGS, XlsFormatterCellMergerNodeModel.DEFAULT_ALL_TAGS);
		DialogComponentBoolean alltagsComponent = new DialogComponentBoolean(alltags,"applies to all tags");
		this.addDialogComponent(alltagsComponent);
		this.setHorizontalPlacement(false);
		
		alltags.addChangeListener(new ChangeListener(){
			public void stateChanged(final ChangeEvent e) {
				tag.setEnabled(!alltags.getBooleanValue());
			}
		});
	}
	
	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);
		
		UiValidation.validateTagField(tag);
	}
}