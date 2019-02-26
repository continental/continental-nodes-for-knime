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

package com.continental.knime.xlsformatter.rowcolumnsizer;

import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.knime.core.data.DataColumnSpec;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DoubleValue;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentButtonGroup;
import org.knime.core.node.defaultnodesettings.DialogComponentNumber;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.DialogComponentStringSelection;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelDouble;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObjectSpec;

import com.continental.knime.xlsformatter.commons.UiValidation;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;
import com.continental.knime.xlsformatter.rowcolumnsizer.XlsRowColumnSizerNodeModel.ControlTableStyle;
import com.continental.knime.xlsformatter.rowcolumnsizer.XlsRowColumnSizerNodeModel.DimensionToSize;

public class XlsRowColumnSizerNodeDialog extends DefaultNodeSettingsPane {

	private static final String[] CONTROLTABLESTYLE_ARRAY = XlsFormatterUiOptions.getDropdownArrayFromEnum(ControlTableStyle.values());
	private static final String[] DIMENSIONTOSIZE_ARRAY = XlsFormatterUiOptions.getDropdownArrayFromEnum(DimensionToSize.values());

	protected boolean hasDoubleControlTableInput = false; // will be set with last spec info in overwritten loadAdditionalSettingsFrom method
	
	SettingsModelString controlTableStyle;
	SettingsModelString dimensionToSize;
	SettingsModelString tag;
	SettingsModelDouble size;
	SettingsModelBoolean autoSize;
	
	protected XlsRowColumnSizerNodeDialog() {
		super();

		this.createNewGroup(XlsFormatterUiOptions.UI_LABEL_CONTROLTABLESTYLE_GROUPTITLE + " (automatically set based on the provided control table)");
		setHorizontalPlacement(true);
		controlTableStyle = new SettingsModelString(XlsRowColumnSizerNodeModel.CFGKEY_CONTROLTABLESTYLE,
				hasDoubleControlTableInput ? ControlTableStyle.DIRECT.toString() : ControlTableStyle.STANDARD.toString());
		DialogComponentButtonGroup controlTableStyleComponent = new DialogComponentButtonGroup( //ControlTableStyle
				controlTableStyle, "", false, CONTROLTABLESTYLE_ARRAY, CONTROLTABLESTYLE_ARRAY);
		controlTableStyleComponent.setToolTipText("The control table style is set automatically depending on the connected control table. Connect either a standard, tag-based control table or one containing the direct row or column sizes in Integer-typed columns.");
		this.addDialogComponent(controlTableStyleComponent);
		setHorizontalPlacement(false);   


		this.createNewGroup("Row and Column Size");
		dimensionToSize = new SettingsModelString(XlsRowColumnSizerNodeModel.CFGKEY_ROW_COLUMN_SIZE, XlsRowColumnSizerNodeModel.DEFAULT_ROW_COLUMN_SIZE);
		DialogComponentStringSelection rowColumnSizeComponent = new DialogComponentStringSelection( //ControlTableStyle
				dimensionToSize, "change", DIMENSIONTOSIZE_ARRAY);
		rowColumnSizeComponent.setToolTipText("Select the function you want to apply in your Xls Sheet.");
		this.addDialogComponent(rowColumnSizeComponent);

		this.createNewGroup("Tag and Size");
		setHorizontalPlacement(true);
		tag = new SettingsModelString(XlsRowColumnSizerNodeModel.CFGKEY_TAGSTRING,XlsRowColumnSizerNodeModel.DEFAULT_TAGSTRING);
		DialogComponentString tagstringComponent = new DialogComponentString(tag, XlsFormatterUiOptions.UI_LABEL_SINGLE_TAG); 
		tagstringComponent.setToolTipText(XlsFormatterUiOptions.UI_TOOLTIP_SINGLE_TAG);
		this.addDialogComponent(tagstringComponent);
		
		size = new SettingsModelDouble(XlsRowColumnSizerNodeModel.CFGKEY_SIZE, XlsRowColumnSizerNodeModel.DEFAULT_SIZE);
		this.addDialogComponent(new DialogComponentNumber( //Size
				size, "size", 1));
		
		autoSize = new SettingsModelBoolean(XlsRowColumnSizerNodeModel.CFGKEY_AUTO_SIZE, XlsRowColumnSizerNodeModel.DEFAULT_AUTO_SIZE);
		DialogComponentBoolean autoSizeComponent = new DialogComponentBoolean(autoSize, "auto-size");
		autoSizeComponent.setToolTipText("Auto-size functionality is only available for columns.");
		this.addDialogComponent(autoSizeComponent);

		size.addChangeListener(new ChangeListener() {
			public void stateChanged(final ChangeEvent e) {
				if (size.isEnabled() && size.getDoubleValue() <= 0)
					size.setDoubleValue(0.1);
			}
		});
		
		dimensionToSize.addChangeListener(new ChangeListener(){
			public void stateChanged(final ChangeEvent e) {
				if (dimensionToSize.isEnabled() && tag.isEnabled()) {
					boolean action = dimensionToSize.getStringValue().equals(DimensionToSize.COLUMN.toString());					
					size.setEnabled(true);
					autoSize.setEnabled(action);
				}
			}
		});
		
		autoSize.addChangeListener(new ChangeListener(){
			public void stateChanged(final ChangeEvent e) {
				if(autoSize.isEnabled()) {
					boolean action = !controlTableStyle.getStringValue().equals(XlsFormatterUiOptions.UI_LABEL_CONTROLTABLESTYLE_STANDARD) ||
							!dimensionToSize.getStringValue().equals(DimensionToSize.COLUMN.toString()) ||
							!autoSize.getBooleanValue();
					size.setEnabled(action);
				}
			}
		});
	}
	
	@Override
	public void loadAdditionalSettingsFrom(NodeSettingsRO settings, PortObjectSpec[] specs)
			throws NotConfigurableException {
		
		super.loadAdditionalSettingsFrom(settings, specs);
		
		DataTableSpec spec = (DataTableSpec)specs[0];
  	if (spec.getNumColumns() == 0)
  		throw new NotConfigurableException("\nThis node cannot be configured without a connected control table at the first inport.");
		DataColumnSpec columnSpec = spec.getColumnSpec(0);
  	hasDoubleControlTableInput = columnSpec.getType().isCompatible(DoubleValue.class);
  	
  	controlTableStyle.setStringValue(hasDoubleControlTableInput ? ControlTableStyle.DIRECT.toString() : ControlTableStyle.STANDARD.toString());
  	controlTableStyle.setEnabled(false);
  	tag.setEnabled(!hasDoubleControlTableInput);
  	autoSize.setEnabled(!hasDoubleControlTableInput);
  	size.setEnabled(!hasDoubleControlTableInput && !autoSize.getBooleanValue());  	
	}
	
	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);
		
		if (!hasDoubleControlTableInput)
			UiValidation.validateTagField(tag);
	}
}
