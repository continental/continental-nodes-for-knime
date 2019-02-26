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

package com.continental.knime.xlsformatter.conditionalformatter;

import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentColorChooser;
import org.knime.core.node.defaultnodesettings.DialogComponentNumber;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelColor;
import org.knime.core.node.defaultnodesettings.SettingsModelDouble;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

import com.continental.knime.xlsformatter.commons.UiValidation;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;

public class XlsFormatterConditionalFormatterNodeDialog extends DefaultNodeSettingsPane {

	SettingsModelString tag;
	SettingsModelBoolean midScalePointActive;
	SettingsModelDouble min;
	SettingsModelDouble mid;
	SettingsModelDouble max;
	
	protected XlsFormatterConditionalFormatterNodeDialog() {
		super();             

		this.createNewGroup("Tag Selection");
		tag = new SettingsModelString(XlsFormatterConditionalFormatterNodeModel.CFGKEY_TAGSTRING, XlsFormatterConditionalFormatterNodeModel.DEFAULT_TAGSTRING);
		this.addDialogComponent(new DialogComponentString(tag, XlsFormatterUiOptions.UI_LABEL_SINGLE_TAG));        


		this.createNewGroup("Conditional Formatting Settings");
		midScalePointActive = new SettingsModelBoolean(XlsFormatterConditionalFormatterNodeModel.CFGKEY_MIDSCALEPOINT_ON,XlsFormatterConditionalFormatterNodeModel.DEFAULT_MIDSCALEPOINT_ON);
		this.addDialogComponent(new DialogComponentBoolean( 
				midScalePointActive, "Mid scale point needed?"));        

		this.setHorizontalPlacement(true);
		min = new SettingsModelDouble(XlsFormatterConditionalFormatterNodeModel.CFGKEY_MIN_VALUE,XlsFormatterConditionalFormatterNodeModel.DEFAULT_MIN_VALUE);
		this.addDialogComponent(new DialogComponentNumber( 
				min, "min", 0.1)); 
		SettingsModelColor mincolor = new SettingsModelColor(XlsFormatterConditionalFormatterNodeModel.CFGKEY_MIN_COLOR,XlsFormatterConditionalFormatterNodeModel.DEFAULT_MIN_COLOR);
		this.addDialogComponent(new DialogComponentColorChooser( 
				mincolor, "min color",true));
		this.setHorizontalPlacement(false);

		this.setHorizontalPlacement(true);
		mid = new SettingsModelDouble(XlsFormatterConditionalFormatterNodeModel.CFGKEY_MID_VALUE,XlsFormatterConditionalFormatterNodeModel.DEFAULT_MID_VALUE);
		this.addDialogComponent(new DialogComponentNumber( 
				mid, "mid", 0.1)); 
		SettingsModelColor midcolor = new SettingsModelColor(XlsFormatterConditionalFormatterNodeModel.CFGKEY_MID_COLOR,XlsFormatterConditionalFormatterNodeModel.DEFAULT_MID_COLOR);
		this.addDialogComponent(new DialogComponentColorChooser( 
				midcolor, "mid color",true));
		this.setHorizontalPlacement(false);

		this.setHorizontalPlacement(true);
		max = new SettingsModelDouble(XlsFormatterConditionalFormatterNodeModel.CFGKEY_MAX_VALUE,XlsFormatterConditionalFormatterNodeModel.DEFAULT_MAX_VALUE);
		this.addDialogComponent(new DialogComponentNumber( 
				max, "max", 0.1)); 
		SettingsModelColor maxcolor = new SettingsModelColor(XlsFormatterConditionalFormatterNodeModel.CFGKEY_MAX_COLOR,XlsFormatterConditionalFormatterNodeModel.DEFAULT_MAX_COLOR);
		this.addDialogComponent(new DialogComponentColorChooser( 
				maxcolor, "max color",true));
		this.setHorizontalPlacement(false);

		midScalePointActive.addChangeListener(new ChangeListener(){
			public void stateChanged(final ChangeEvent e) {
				mid.setEnabled(midScalePointActive.getBooleanValue());
				midcolor.setEnabled(midScalePointActive.getBooleanValue());
			}
		});
	}
	
	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);
		
		UiValidation.validateTagField(tag);
		
		if (min.getDoubleValue() > max.getDoubleValue() ||
				(midScalePointActive.getBooleanValue() && (min.getDoubleValue() > mid.getDoubleValue() || mid.getDoubleValue() > max.getDoubleValue())))
			throw new InvalidSettingsException(midScalePointActive.getBooleanValue() ?
					"Min must be smaller than mid, and mid must be smaller than max value." :
					"Min must be smaller than max value.");
	}
}
