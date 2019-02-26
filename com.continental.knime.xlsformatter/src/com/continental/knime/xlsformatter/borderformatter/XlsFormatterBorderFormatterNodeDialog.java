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

import java.util.List;

import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentColorChooser;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.DialogComponentStringSelection;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelColor;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

import com.continental.knime.xlsformatter.commons.UiValidation;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;

public class XlsFormatterBorderFormatterNodeDialog extends DefaultNodeSettingsPane {

	static List<String> borderStyleSelectionList;

	static {
		borderStyleSelectionList = XlsFormatterUiOptions.getDropdownListFromEnum(XlsFormatterState.BorderStyle.values());
		borderStyleSelectionList.remove(XlsFormatterState.BorderStyle.UNMODIFIED.toString().toLowerCase()); // remove 'unmodified', because in this case it wouldn't make sense to have this node at all
		borderStyleSelectionList.remove(XlsFormatterState.BorderStyle.NONE.toString().toLowerCase()); // remove 'none', because "undo" functionality is currently not offered in the UI
	}

	SettingsModelString tag;
	SettingsModelBoolean allTags;
	SettingsModelBoolean borderChangeColor;
	SettingsModelBoolean borderTop;
	SettingsModelBoolean borderLeft;
	SettingsModelBoolean borderRight;
	SettingsModelBoolean borderBottom;
	SettingsModelBoolean borderInnerVertical;
	SettingsModelBoolean borderInnerHorizontal;
	
	protected XlsFormatterBorderFormatterNodeDialog() {
		super();

		this.createNewGroup("Tag Selection");
		this.setHorizontalPlacement(true);
		tag = new SettingsModelString(XlsFormatterBorderFormatterNodeModel.CFGKEY_TAGSTRING, XlsFormatterBorderFormatterNodeModel.DEFAULT_TAGSTRING);
		DialogComponentString tagstringComponent = new DialogComponentString(tag , XlsFormatterUiOptions.UI_LABEL_SINGLE_TAG);
		this.addDialogComponent(tagstringComponent);
		
		allTags = new SettingsModelBoolean(XlsFormatterBorderFormatterNodeModel.CFGKEY_ALL_TAGS, XlsFormatterBorderFormatterNodeModel.DEFAULT_ALL_TAGS);
		DialogComponentBoolean alltagsComponent = new DialogComponentBoolean(allTags,"applies to all tags");
		this.addDialogComponent(alltagsComponent);
		this.setHorizontalPlacement(false);


		this.createNewGroup("Border Style and Color");
		this.setHorizontalPlacement(true);
		SettingsModelString borderstyle = new SettingsModelString(XlsFormatterBorderFormatterNodeModel.CFGKEY_BORDER_STYLE, XlsFormatterBorderFormatterNodeModel.DEFAULT_BORDER_STYLE);
		this.addDialogComponent(new DialogComponentStringSelection(
				borderstyle, "border style",borderStyleSelectionList));
		this.setHorizontalPlacement(false);
		
		this.setHorizontalPlacement(true);
		borderChangeColor = new SettingsModelBoolean(XlsFormatterBorderFormatterNodeModel.CFGKEY_BORDER_CHANGECOLOR, XlsFormatterBorderFormatterNodeModel.DEFAULT_BORDER_CHANGECOLOR);
		this.addDialogComponent(new DialogComponentBoolean(
				borderChangeColor, "Change border color?"));
		
		SettingsModelColor borderColor = new SettingsModelColor(XlsFormatterBorderFormatterNodeModel.CFGKEY_BORDER_COLOR, XlsFormatterBorderFormatterNodeModel.DEFAULT_BORDER_COLOR);
		this.addDialogComponent(new DialogComponentColorChooser(
				borderColor, "border color",true));
		this.setHorizontalPlacement(false);

		
		this.createNewGroup("Outer Border Settings");
		borderTop = new SettingsModelBoolean(XlsFormatterBorderFormatterNodeModel.CFGKEY_BORDER_TOP, XlsFormatterBorderFormatterNodeModel.DEFAULT_BORDER_TOP);
		this.addDialogComponent(new DialogComponentBoolean(
				borderTop, "top"));
		this.setHorizontalPlacement(true);
		
		borderLeft = new SettingsModelBoolean(XlsFormatterBorderFormatterNodeModel.CFGKEY_BORDER_LEFT, XlsFormatterBorderFormatterNodeModel.DEFAULT_BORDER_LEFT);
		this.addDialogComponent(new DialogComponentBoolean(
				borderLeft, "left"));
		
		borderRight = new SettingsModelBoolean(XlsFormatterBorderFormatterNodeModel.CFGKEY_BORDER_RIGHT, XlsFormatterBorderFormatterNodeModel.DEFAULT_BORDER_RIGHT);
		this.addDialogComponent(new DialogComponentBoolean(
				borderRight, "right"));
		this.setHorizontalPlacement(false);
		
		borderBottom = new SettingsModelBoolean(XlsFormatterBorderFormatterNodeModel.CFGKEY_BORDER_BOTTOM, XlsFormatterBorderFormatterNodeModel.DEFAULT_BORDER_BOTTOM);
		this.addDialogComponent(new DialogComponentBoolean(
				borderBottom, "bottom"));

		
		this.createNewGroup("Inner Border Settings");
		this.setHorizontalPlacement(true);
		borderInnerVertical = new SettingsModelBoolean(XlsFormatterBorderFormatterNodeModel.CFGKEY_BORDER_INNER_VERTICAL, XlsFormatterBorderFormatterNodeModel.DEFAULT_BORDER_INNER_VERTICAL);
		this.addDialogComponent(new DialogComponentBoolean(
				borderInnerVertical , "inner vertical"));
		
		borderInnerHorizontal = new SettingsModelBoolean(XlsFormatterBorderFormatterNodeModel.CFGKEY_BORDER_INNER_HORIZONTAL, XlsFormatterBorderFormatterNodeModel.DEFAULT_BORDER_INNER_HORIZONTAL);
		this.addDialogComponent(new DialogComponentBoolean(
				borderInnerHorizontal, "inner horizontal"));
		this.setHorizontalPlacement(false);

		borderChangeColor.addChangeListener(new ChangeListener(){
			public void stateChanged(final ChangeEvent e) {
				borderColor.setEnabled(borderChangeColor.getBooleanValue());
			}
		});

		allTags.addChangeListener(new ChangeListener(){
			public void stateChanged(final ChangeEvent e) {
				tag.setEnabled(!allTags.getBooleanValue());
			}
		});
	}
	
	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);
		
		if (!allTags.getBooleanValue())
			UiValidation.validateTagField(tag, "Alternatively, activate the 'all tags' option to apply the border to all different tags found in the control table.");

		if (!borderTop.getBooleanValue() && !borderLeft.getBooleanValue() && !borderRight.getBooleanValue() &&
				!borderBottom.getBooleanValue() && !borderInnerVertical.getBooleanValue() && !borderInnerHorizontal.getBooleanValue())
			throw new InvalidSettingsException("At least one border option needs to be activated, otherwise this node would be of no effect.");
	}
}

