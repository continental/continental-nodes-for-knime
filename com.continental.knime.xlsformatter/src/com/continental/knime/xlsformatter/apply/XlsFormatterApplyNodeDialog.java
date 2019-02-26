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

package com.continental.knime.xlsformatter.apply;

import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentFileChooser;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;


public class XlsFormatterApplyNodeDialog extends DefaultNodeSettingsPane {

	protected XlsFormatterApplyNodeDialog() {
		super();

		DialogComponentFileChooser inputComponent = new DialogComponentFileChooser( // Select Input File
				new SettingsModelString(XlsFormatterApplyNodeModel.CFGKEY_INPUTFILE, XlsFormatterApplyNodeModel.DEFAULT_INPUTFILE), "Input File");
		inputComponent.setToolTipText("Select the .xlsx file to format.");
		inputComponent.setBorderTitle("Input File");
		inputComponent.setAllowSystemPropertySubstitution(true);
		this.addDialogComponent(inputComponent);

		DialogComponentFileChooser destinationComponent = new DialogComponentFileChooser( // Select Output File Destination
				new SettingsModelString(XlsFormatterApplyNodeModel.CFGKEY_OUTPUTFILE, XlsFormatterApplyNodeModel.DEFAULT_OUTPUTFILE), "Output File",1);
		destinationComponent.setToolTipText("Select the path name to store your formatted .xlsx file.");
		destinationComponent.setBorderTitle("Output File");
		destinationComponent.setAllowSystemPropertySubstitution(true);

		this.addDialogComponent(destinationComponent);

		this.setHorizontalPlacement(true);
		this.addDialogComponent(new DialogComponentBoolean(
				new SettingsModelBoolean(XlsFormatterApplyNodeModel.CFGKEY_OVERWRITE,XlsFormatterApplyNodeModel.DEFAULT_OVERWRITE),
				"overwrite output file"));
		this.addDialogComponent(new DialogComponentBoolean(
				new SettingsModelBoolean(XlsFormatterApplyNodeModel.CFGKEY_OPENOUTPUTFILE,XlsFormatterApplyNodeModel.DEFAULT_OPENOUTPUTFILE),
				"open output file after execution"));
	}
}

