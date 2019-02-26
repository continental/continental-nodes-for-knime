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

package com.continental.knime.xlsformatter.cellbackgroundcolorizer;

import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentButtonGroup;
import org.knime.core.node.defaultnodesettings.DialogComponentColorChooser;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.DialogComponentStringSelection;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelColor;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

import com.continental.knime.xlsformatter.commons.UiValidation;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;

public class XlsFormatterCellBackgroundColorizerNodeDialog extends DefaultNodeSettingsPane {

	SettingsModelString controlTableStyle;
	SettingsModelString tag;
	SettingsModelBoolean backgroundColor;
	SettingsModelString backgroundPattern;

	protected XlsFormatterCellBackgroundColorizerNodeDialog() {
		super();

		this.createNewGroup(XlsFormatterUiOptions.UI_LABEL_CONTROLTABLESTYLE_GROUPTITLE);
		setHorizontalPlacement(true);
		controlTableStyle = new SettingsModelString(XlsFormatterCellBackgroundColorizerNodeModel.CFGKEY_CONTROLTABLESTYLE, XlsFormatterCellBackgroundColorizerNodeModel.DEFAULT_CONTROLTABLESTYLE);
		String[] controlTableStyleButtonArray = { XlsFormatterCellBackgroundColorizerNodeModel.OPTION_CONTROLTABLESTYLE_STANDARD, XlsFormatterCellBackgroundColorizerNodeModel.OPTION_CONTROLTABLESTYLE_DIRECT};
		DialogComponentButtonGroup controlTableStyleComponent = new DialogComponentButtonGroup(
				controlTableStyle, "", false, controlTableStyleButtonArray, controlTableStyleButtonArray);
		controlTableStyleComponent.setToolTipText("Either use a tag to define your background color or use RGB color-codes in your control table (e.g. 255/0/255 or #FF00FF).");
		this.addDialogComponent(controlTableStyleComponent);
		setHorizontalPlacement(false);        

		
		this.createNewGroup("Tag Selection");
		tag = new SettingsModelString(XlsFormatterCellBackgroundColorizerNodeModel.CFGKEY_TAGSTRING,XlsFormatterCellBackgroundColorizerNodeModel.DEFAULT_TAGSTRING);
		DialogComponentString tagstringComponent = new DialogComponentString(tag, XlsFormatterUiOptions.UI_LABEL_SINGLE_TAG); 
		tagstringComponent.setToolTipText(XlsFormatterUiOptions.UI_TOOLTIP_SINGLE_TAG);
		this.addDialogComponent(tagstringComponent); 

		
		this.createNewGroup("Background Color");
		this.setHorizontalPlacement(true);
		backgroundColor = new SettingsModelBoolean(XlsFormatterCellBackgroundColorizerNodeModel.CFGKEY_BACKGROUND_COLOR_SELECTION,XlsFormatterCellBackgroundColorizerNodeModel.DEFAULT_BACKGROUND_COLOR_SELECTION);
		DialogComponentBoolean backgroundcolorselectionComponent = new DialogComponentBoolean(
				backgroundColor, "Change color?");
		backgroundcolorselectionComponent.setToolTipText("Do not modify, remove or select the background color.");
		this.addDialogComponent(backgroundcolorselectionComponent);
		
		SettingsModelColor backgroundcolor = new SettingsModelColor(XlsFormatterCellBackgroundColorizerNodeModel.CFGKEY_BACKGROUNDCOLOR, XlsFormatterCellBackgroundColorizerNodeModel.DEFAULT_BACKGROUNDCOLOR);
		this.addDialogComponent(new DialogComponentColorChooser(
				backgroundcolor, "color", true));
		this.setHorizontalPlacement(false);

		
		this.createNewGroup("Pattern Fill");
		backgroundPattern = new SettingsModelString(XlsFormatterCellBackgroundColorizerNodeModel.CFGKEY_BACKGROUND_PATTERN_SELECTION,XlsFormatterCellBackgroundColorizerNodeModel.DEFAULT_BACKGROUND_PATTERN_SELECTION);
		DialogComponentStringSelection backgroundpatternselectionComponent = new DialogComponentStringSelection(
				backgroundPattern, "pattern fill", XlsFormatterCellBackgroundColorizerNodeModel.fillPatternDropdownOptions);
		backgroundpatternselectionComponent.setToolTipText("Do not modify a previously defined pattern fill or select a pattern now.");
		this.addDialogComponent(backgroundpatternselectionComponent);
		
		this.setHorizontalPlacement(true);
		SettingsModelBoolean backgroundpatterncolorselection = new SettingsModelBoolean(XlsFormatterCellBackgroundColorizerNodeModel.CFGKEY_BACKGROUND_PATTERN_COLOR_SELECTION, XlsFormatterCellBackgroundColorizerNodeModel.DEFAULT_BACKGROUND_PATTERN_COLOR_SELECTION);
		this.addDialogComponent(new DialogComponentBoolean(
				backgroundpatterncolorselection, "Change color?"));
		
		SettingsModelColor backgroundpatterncolor = new SettingsModelColor(XlsFormatterCellBackgroundColorizerNodeModel.CFGKEY_BACKGROUND_PATTERN_COLOR, XlsFormatterCellBackgroundColorizerNodeModel.DEFAULT_BACKGROUND_PATTERN_COLOR);
		this.addDialogComponent(new DialogComponentColorChooser(
				backgroundpatterncolor, "color", true));


		controlTableStyle.addChangeListener(new ChangeListener(){
			public void stateChanged(final ChangeEvent e) {
				boolean action = controlTableStyle.getStringValue().equals(controlTableStyleButtonArray[0]);
				tag.setEnabled(action);
				backgroundColor.setEnabled(action);
				backgroundPattern.setEnabled(action);
				if(action) {
					if(backgroundColor.getBooleanValue())
						backgroundcolor.setEnabled(action);
					if(backgroundPattern.getStringValue().equals(XlsFormatterState.FormattingFlag.ON.toString().toLowerCase())) {
						backgroundpatterncolorselection.setEnabled(action);
						backgroundpatterncolor.setEnabled(action);
					}
				}
				else {
					backgroundpatterncolorselection.setEnabled(action);
					backgroundpatterncolor.setEnabled(action);
					backgroundcolor.setEnabled(action);
				}
			}
		});

		backgroundColor.addChangeListener(new ChangeListener() {
			public void stateChanged(final ChangeEvent e) {
				boolean action = backgroundColor.getBooleanValue();
				backgroundcolor.setEnabled(action);
			}
		});

		backgroundPattern.addChangeListener(new ChangeListener() {
			public void stateChanged(final ChangeEvent e) {
				if(backgroundPattern.isEnabled() &&
						!backgroundPattern.getStringValue().equals(XlsFormatterState.FillPattern.NONE.toString().toLowerCase()))
				backgroundpatterncolorselection.setEnabled(true);
				if(backgroundpatterncolorselection.isEnabled() && 
						backgroundpatterncolorselection.getBooleanValue())
					backgroundpatterncolor.setEnabled(true);
			}
		});

		backgroundpatterncolorselection.addChangeListener(new ChangeListener() {
			public void stateChanged(final ChangeEvent e) {
				boolean action = backgroundpatterncolorselection.getBooleanValue();
				backgroundpatterncolor.setEnabled(action);
			}
		});
	}

	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);

		if (controlTableStyle.getStringValue().equals(XlsFormatterCellBackgroundColorizerNodeModel.OPTION_CONTROLTABLESTYLE_STANDARD)) {
			UiValidation.validateTagField(tag);
			if (!backgroundColor.getBooleanValue() && backgroundPattern.getStringValue().equals(XlsFormatterState.FillPattern.UNMODIFIED.toString().toLowerCase()))
				throw new InvalidSettingsException("Background color or fill pattern need to be changed, otherwise this node would not have any effect.");
		}
	}
}

