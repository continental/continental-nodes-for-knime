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

package com.continental.knime.xlsformatter.fontformatter;

import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentColorChooser;
import org.knime.core.node.defaultnodesettings.DialogComponentNumber;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelColor;
import org.knime.core.node.defaultnodesettings.SettingsModelInteger;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObjectSpec;

import com.continental.knime.xlsformatter.commons.UiValidation;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;

public class XlsFormatterFontFormatterNodeDialog extends DefaultNodeSettingsPane {

	SettingsModelString tag;
	SettingsModelBoolean bold;
	SettingsModelBoolean italic;
	SettingsModelBoolean underline;
	SettingsModelBoolean changeFontSize;
	SettingsModelBoolean changeColor;
	SettingsModelColor fontcolor;
	SettingsModelInteger size;
	
	ChangeListener changeListener = new FontDialogChangeListener();
	
	protected XlsFormatterFontFormatterNodeDialog() {
		super();

		this.createNewGroup("Tag Selection");
		tag = new SettingsModelString(XlsFormatterFontFormatterNodeModel.CFGKEY_TAG, XlsFormatterFontFormatterNodeModel.DEFAULT_TAG);
		this.addDialogComponent(new DialogComponentString(tag, XlsFormatterUiOptions.UI_LABEL_SINGLE_TAG, true, 10));        


		this.createNewGroup("Font Specification");        
		setHorizontalPlacement(true);
		bold = new SettingsModelBoolean(XlsFormatterFontFormatterNodeModel.CFGKEY_BOLD, XlsFormatterFontFormatterNodeModel.DEFAULT_BOLD);
		this.addDialogComponent(new DialogComponentBoolean(bold, "bold"));
		
		italic = new SettingsModelBoolean(XlsFormatterFontFormatterNodeModel.CFGKEY_ITALIC, XlsFormatterFontFormatterNodeModel.DEFAULT_ITALIC);
		this.addDialogComponent(new DialogComponentBoolean(italic, "italic"));
		
		underline = new SettingsModelBoolean(XlsFormatterFontFormatterNodeModel.CFGKEY_UNDERLINE, XlsFormatterFontFormatterNodeModel.DEFAULT_UNDERLINE);
		this.addDialogComponent(new DialogComponentBoolean(underline, "underline"));

		setHorizontalPlacement(false);
		setHorizontalPlacement(true);
		changeFontSize = new SettingsModelBoolean(XlsFormatterFontFormatterNodeModel.CFGKEY_CHANGESIZE, XlsFormatterFontFormatterNodeModel.DEFAULT_CHANGESIZE);
		this.addDialogComponent(new DialogComponentBoolean(changeFontSize, "Change font size?"));
		
		size = new SettingsModelInteger(XlsFormatterFontFormatterNodeModel.CFGKEY_SIZE, XlsFormatterFontFormatterNodeModel.DEFAULT_SIZE);
		size.setEnabled(XlsFormatterFontFormatterNodeModel.DEFAULT_CHANGESIZE);
		this.addDialogComponent(new DialogComponentNumber( // Size
				size, "font size", 1));

		setHorizontalPlacement(false);
		setHorizontalPlacement(true);
		changeColor = new SettingsModelBoolean(XlsFormatterFontFormatterNodeModel.CFGKEY_CHANGECOLOR, XlsFormatterFontFormatterNodeModel.DEFAULT_CHANGECOLOR);
		this.addDialogComponent(new DialogComponentBoolean(changeColor, "Change color?"));
		
		fontcolor = new SettingsModelColor(XlsFormatterFontFormatterNodeModel.CFGKEY_FONTCOLOR, XlsFormatterFontFormatterNodeModel.DEFAULT_FONTCOLOR);
		fontcolor.setEnabled(XlsFormatterFontFormatterNodeModel.DEFAULT_CHANGECOLOR);
		this.addDialogComponent(new DialogComponentColorChooser( // color
				fontcolor, "color", true));
		
		changeColor.addChangeListener(changeListener);
		size.addChangeListener(changeListener);
		changeFontSize.addChangeListener(changeListener);
	}
	
	class FontDialogChangeListener implements ChangeListener {
		public void stateChanged(final ChangeEvent e) {
			fontcolor.setEnabled(changeColor.getBooleanValue());
			size.setEnabled(changeFontSize.getBooleanValue());
			if (size.getIntValue() <= 0)
				size.setIntValue(1);
		}
	}	
	
	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);
		
		UiValidation.validateTagField(tag);
		
		if (!bold.getBooleanValue() && !italic.getBooleanValue() && !underline.getBooleanValue() && !changeColor.getBooleanValue() &&
				!changeFontSize.getBooleanValue())
			throw new InvalidSettingsException("One of the dialog's options need to be activate, otherwise this node would not have any effect.");
	}
	
	@Override
	public void loadAdditionalSettingsFrom(NodeSettingsRO settings, PortObjectSpec[] specs)
			throws NotConfigurableException {
		super.loadAdditionalSettingsFrom(settings, specs);
		changeListener.stateChanged(null);
	}
}
