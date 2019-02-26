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

package com.continental.knime.xlsformatter.cellformatter;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentString;
import org.knime.core.node.defaultnodesettings.DialogComponentStringSelection;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

import com.continental.knime.xlsformatter.commons.UiValidation;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;

public class XlsFormatterCellFormatterNodeDialog extends DefaultNodeSettingsPane {

	static List<String> flagSelectionList;
	static List<String> horizontalAlignmentSelectionList;
	static List<String> verticalAlignmentSelectionList;
	static List<String> textRotationSelectionList;
	static List<String> dataTypeSelectionList;
	static List<String> textPresetSelectionList;

	static {
		flagSelectionList = XlsFormatterUiOptions.getDropdownListFromEnum(XlsFormatterState.FormattingFlag.values());
		horizontalAlignmentSelectionList = XlsFormatterUiOptions.getDropdownListFromEnum(XlsFormatterState.CellAlignmentHorizontal.values());
		verticalAlignmentSelectionList = XlsFormatterUiOptions.getDropdownListFromEnum(XlsFormatterState.CellAlignmentVertical.values());
		dataTypeSelectionList = Arrays.stream(XlsFormatterState.CellDataType.values())
				.filter(v -> v != XlsFormatterState.CellDataType.FORMULA).map(x -> x.toString()).collect(Collectors.toList());
		textPresetSelectionList = Arrays.stream(XlsFormatterCellFormatterNodeModel.TextPresets.values())
				.map(x -> x.toString()).collect(Collectors.toList());
		textRotationSelectionList = new ArrayList<String>(6);
		textRotationSelectionList.add(XlsFormatterState.FormattingFlag.UNMODIFIED.toString().toLowerCase());
		textRotationSelectionList.add("+90°");
		textRotationSelectionList.add("+45°");
		textRotationSelectionList.add("0°");
		textRotationSelectionList.add("-45°");
		textRotationSelectionList.add("-90°");
	}

	SettingsModelString tag;
	SettingsModelString horizontalAlignment;
	SettingsModelString verticalAlignment;
	SettingsModelString textRotationAngle;
	SettingsModelBoolean wordWrap;
	SettingsModelString cellStyleConversion;
	SettingsModelString textFormat;

	protected XlsFormatterCellFormatterNodeDialog() {
		super();

		this.createNewGroup("Tag Selection");
		tag = new SettingsModelString(XlsFormatterCellFormatterNodeModel.CFGKEY_TAGSTRING,XlsFormatterCellFormatterNodeModel.DEFAULT_TAGSTRING);
		this.addDialogComponent(new DialogComponentString(tag, "applies to tag (single tag only)"));        


		this.createNewGroup("Text Position and Layout");        
		setHorizontalPlacement(false);
		horizontalAlignment = new SettingsModelString(XlsFormatterCellFormatterNodeModel.CFGKEY_HORIZONTALALIGNMENT,XlsFormatterCellFormatterNodeModel.DEFAULT_HORIZONTALALIGNMENT);
		this.addDialogComponent(new DialogComponentStringSelection(
				horizontalAlignment, "horizontal alignment", horizontalAlignmentSelectionList));

		verticalAlignment = new SettingsModelString(XlsFormatterCellFormatterNodeModel.CFGKEY_VERTICALALIGNMENT,XlsFormatterCellFormatterNodeModel.DEFAULT_VERTICALALIGNMENT);
		this.addDialogComponent(new DialogComponentStringSelection(
				verticalAlignment, "vertical alignment", verticalAlignmentSelectionList));

		textRotationAngle = new SettingsModelString(XlsFormatterCellFormatterNodeModel.CFGKEY_TEXTROTATION,XlsFormatterCellFormatterNodeModel.DEFAULT_TEXTROTATION);
		this.addDialogComponent(new DialogComponentStringSelection(
				textRotationAngle, "text rotation angle", textRotationSelectionList));

		wordWrap = new SettingsModelBoolean(XlsFormatterCellFormatterNodeModel.CFGKEY_WORDWRAP,XlsFormatterCellFormatterNodeModel.DEFAULT_WORDWRAP);
		this.addDialogComponent(new DialogComponentBoolean(
				wordWrap,	"word wrap"));


		this.createNewGroup("Data Type and Format");
		cellStyleConversion = new SettingsModelString(XlsFormatterCellFormatterNodeModel.CFGKEY_CELL_STYLE, XlsFormatterCellFormatterNodeModel.DEFAULT_CELL_STYLE);
		this.addDialogComponent(new DialogComponentStringSelection(
				cellStyleConversion, "cell style conversion (from String to)", dataTypeSelectionList));

		this.setHorizontalPlacement(true);
		textFormat = new SettingsModelString(XlsFormatterCellFormatterNodeModel.CFGKEY_TEXT_FORMAT,XlsFormatterCellFormatterNodeModel.DEFAULT_TEXT_FORMAT);
		DialogComponentString textFormatComponent = new DialogComponentString(textFormat , "text format");
		textFormatComponent.setToolTipText("Set text format to e.g. default, percent (0.00%), whole number (#,##0), ...");
		this.addDialogComponent(textFormatComponent);
		SettingsModelString textPresets = new SettingsModelString(XlsFormatterCellFormatterNodeModel.CFGKEY_TEXT_PRESETS , XlsFormatterCellFormatterNodeModel.DEFAULT_TEXT_PRESETS);
		DialogComponentStringSelection textPresetsComponent = new DialogComponentStringSelection(textPresets , "" , textPresetSelectionList);
		this.addDialogComponent(textPresetsComponent);

		textPresets.addChangeListener(new ChangeListener(){
			public void stateChanged(final ChangeEvent e) {
				if(!textPresets.getStringValue().equals(XlsFormatterCellFormatterNodeModel.TextPresets.UNMODIFIED.toString())) {
					if(textPresets.getStringValue().equals(XlsFormatterCellFormatterNodeModel.TextPresets.PERCENT.toString()))
						textFormat.setStringValue(XlsFormatterCellFormatterNodeModel.TextPresets.PERCENT.getTextFormat());
					if(textPresets.getStringValue().equals(XlsFormatterCellFormatterNodeModel.TextPresets.INTEGER.toString()))
						textFormat.setStringValue(XlsFormatterCellFormatterNodeModel.TextPresets.INTEGER.getTextFormat());
					if(textPresets.getStringValue().equals(XlsFormatterCellFormatterNodeModel.TextPresets.THOUSANDSEPARATED.toString()))
						textFormat.setStringValue(XlsFormatterCellFormatterNodeModel.TextPresets.THOUSANDSEPARATED.getTextFormat());
					if(textPresets.getStringValue().equals(XlsFormatterCellFormatterNodeModel.TextPresets.FINANCIAL.toString()))
						textFormat.setStringValue(XlsFormatterCellFormatterNodeModel.TextPresets.FINANCIAL.getTextFormat());
					if(textPresets.getStringValue().equals(XlsFormatterCellFormatterNodeModel.TextPresets.DATE.toString()))
						textFormat.setStringValue(XlsFormatterCellFormatterNodeModel.TextPresets.DATE.getTextFormat());
					if(textPresets.getStringValue().equals(XlsFormatterCellFormatterNodeModel.TextPresets.DATETIME.toString()))
						textFormat.setStringValue(XlsFormatterCellFormatterNodeModel.TextPresets.DATETIME.getTextFormat());
					textPresets.setStringValue(XlsFormatterCellFormatterNodeModel.TextPresets.UNMODIFIED.toString());
				}
			}
		});
	}

	@Override
	public void saveAdditionalSettingsTo(NodeSettingsWO settings) throws InvalidSettingsException {
		super.saveAdditionalSettingsTo(settings);

		UiValidation.validateTagField(tag);

		if (horizontalAlignment.getStringValue().equals(XlsFormatterState.FormattingFlag.UNMODIFIED.toString().toLowerCase()) &&
				verticalAlignment.getStringValue().equals(XlsFormatterState.FormattingFlag.UNMODIFIED.toString().toLowerCase()) &&
				textRotationAngle.getStringValue().equals(XlsFormatterState.FormattingFlag.UNMODIFIED.toString().toLowerCase()) &&
				!wordWrap.getBooleanValue() &&
				cellStyleConversion.getStringValue().equals(XlsFormatterState.CellDataType.UNMODIFIED.toString()) &&
				textFormat.getStringValue().trim().equals(""))
			throw new InvalidSettingsException("At least one setting needs to be activated in this dialog, otherwise this node would not have any effect.");
	}
}
