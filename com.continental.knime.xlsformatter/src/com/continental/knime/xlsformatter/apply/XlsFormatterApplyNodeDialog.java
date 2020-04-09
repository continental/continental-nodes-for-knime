/*
 * Continental Nodes for KNIME
 * Copyright (C) 2020  Continental AG, Hanover, Germany
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

import javax.swing.JFileChooser;
import javax.swing.event.ChangeEvent;
import javax.swing.event.ChangeListener;

import org.knime.core.node.FlowVariableModel;
import org.knime.core.node.defaultnodesettings.DefaultNodeSettingsPane;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.defaultnodesettings.DialogComponentFileChooser;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;


public class XlsFormatterApplyNodeDialog extends DefaultNodeSettingsPane {

	SettingsModelString inputFile;	
	FlowVariableModel flowVarInputFile;
	SettingsModelString outputFile;	
	FlowVariableModel flowVarOutputFile;
	
	protected XlsFormatterApplyNodeDialog() {
		super();
		
		ChangeListener changeListener = new ApplyDialogChangeListener();
		
		// input file dialog:
		inputFile = new SettingsModelString(XlsFormatterApplyNodeModel.CFGKEY_INPUTFILE, XlsFormatterApplyNodeModel.DEFAULT_INPUTFILE);
		flowVarInputFile = createFlowVariableModel(inputFile);
		flowVarInputFile.addChangeListener(changeListener);
		
		DialogComponentFileChooser inputComponent = new DialogComponentFileChooser(inputFile, "Input File", JFileChooser.OPEN_DIALOG, false, flowVarInputFile);
		inputComponent.setToolTipText("Select the .xlsx file to format.");
		inputComponent.setBorderTitle("Input File");
		inputComponent.setAllowSystemPropertySubstitution(true);
		this.addDialogComponent(inputComponent);

		// output/destination file dialog:
		outputFile = new SettingsModelString(XlsFormatterApplyNodeModel.CFGKEY_OUTPUTFILE, XlsFormatterApplyNodeModel.DEFAULT_OUTPUTFILE);
		flowVarOutputFile = createFlowVariableModel(outputFile);
		flowVarOutputFile.addChangeListener(changeListener);
		
		DialogComponentFileChooser destinationComponent = new DialogComponentFileChooser(outputFile, "Output File", JFileChooser.SAVE_DIALOG, false, flowVarOutputFile);
		destinationComponent.setToolTipText("Select the path name to store your formatted .xlsx file.");
		destinationComponent.setBorderTitle("Output File");
		destinationComponent.setAllowSystemPropertySubstitution(true);
		this.addDialogComponent(destinationComponent);

		// other controls:
		this.setHorizontalPlacement(true);
		this.addDialogComponent(new DialogComponentBoolean(
				new SettingsModelBoolean(XlsFormatterApplyNodeModel.CFGKEY_OVERWRITE,XlsFormatterApplyNodeModel.DEFAULT_OVERWRITE),
				"overwrite output file"));
		this.addDialogComponent(new DialogComponentBoolean(
				new SettingsModelBoolean(XlsFormatterApplyNodeModel.CFGKEY_OPENOUTPUTFILE,XlsFormatterApplyNodeModel.DEFAULT_OPENOUTPUTFILE),
				"open output file after execution"));
	}
	
	class ApplyDialogChangeListener implements ChangeListener {
		@SuppressWarnings("deprecation")
		public void stateChanged(final ChangeEvent e) {
			if (e.getSource() == flowVarInputFile) {
				if (flowVarInputFile.isVariableReplacementEnabled())
					inputFile.setStringValue(getAvailableFlowVariables().get(flowVarInputFile.getInputVariableName()).getStringValue()); // since KNIME 4.1, this should be getAvailableFlowVariables(VariableType.StringType.INSTANCE). Leaving the deprecated version for extension compatibility with older versions.
				else
					inputFile.setStringValue(XlsFormatterApplyNodeModel.DEFAULT_INPUTFILE);
			} else if (e.getSource() == flowVarOutputFile) {
				if (flowVarOutputFile.isVariableReplacementEnabled())
					outputFile.setStringValue(getAvailableFlowVariables().get(flowVarOutputFile.getInputVariableName()).getStringValue()); // since KNIME 4.1, this should be getAvailableFlowVariables(VariableType.StringType.INSTANCE). Leaving the deprecated version for extension compatibility with older versions.
				else
					outputFile.setStringValue(XlsFormatterApplyNodeModel.DEFAULT_OUTPUTFILE);
			}
		}
	}
}

