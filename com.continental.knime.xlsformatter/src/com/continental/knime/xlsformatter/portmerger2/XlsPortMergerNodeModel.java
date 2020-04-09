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

package com.continental.knime.xlsformatter.portmerger2;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;

import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.apply.XlsFormatterApplyNodeModel;
import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.commons.XlsFormattingStateValidator;
import com.continental.knime.xlsformatter.commons.XlsFormattingStateValidator.ValidationModes;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterStateMerger;
import com.continental.knime.xlsformatter.porttype.XlsFormatterStateSpec;

public class XlsPortMergerNodeModel extends NodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsFormatterApplyNodeModel.class);
	
	/**
	 * Constructor for the node model.
	 */
	public XlsPortMergerNodeModel() {
		this(2);		
	}
	
	/**
	 * Constructor for the node model.
	 */
	XlsPortMergerNodeModel(final int nrIns) {
		super(getInPortTypes(nrIns), new PortType[] { XlsFormatterState.TYPE });
	}

	XlsPortMergerNodeModel(final PortsConfiguration portsConfiguration) {
    super(portsConfiguration.getInputPorts(), portsConfiguration.getOutputPorts());
	}
	
	private static final PortType[] getInPortTypes(final int nrIns) {
    if (nrIns < 2) {
        throw new IllegalArgumentException("Invalid number of input ports (" + nrIns + "). Merge operation requires at least 2 inputs.");
    }
    PortType[] result = new PortType[nrIns];
    Arrays.fill(result, XlsFormatterState.TYPE_OPTIONAL);
    result[0] = XlsFormatterState.TYPE;
    result[1] = XlsFormatterState.TYPE;
    return result;
	}
	
	/**
	 * {@inheritDoc}
	 */
	protected PortObject[] execute(PortObject[] inObjects, final ExecutionContext exec) throws Exception { 

		XlsFormatterState master = XlsFormatterState.getDeepClone(inObjects[0]); // the master port state that will be added on
		for (int i = 1; i < inObjects.length; i++)
			XlsFormatterStateMerger.mergeFormatterStates(master, (XlsFormatterState)inObjects[i], exec, logger);
		
		WarningMessageContainer warningMessageContainer = new WarningMessageContainer();
		
		if (master.isEmpty())
			warningMessageContainer.addMessage("All inputs were empty XLS Formatting states, so is the generated output.");
		
		XlsFormattingStateValidator.validateState(master, ValidationModes.EVERYTHING, warningMessageContainer, exec, logger);
		
		if (warningMessageContainer.hasMessage())
			setWarningMessage(warningMessageContainer.getMessage());
		
		return new PortObject[] { master };
	}

	
	@Override
	protected PortObjectSpec[] configure(final PortObjectSpec[] inSpecs)
			throws InvalidSettingsException {   	

		for (int i = 0; i < inSpecs.length - 1; i++)
			for (int j = i + 1; j < inSpecs.length; j++)
				if (inSpecs[i] == inSpecs[j])
					throw new InvalidSettingsException("Self-merge is not allowed. Please use different XLS Formatting nodes' result ports as inputs to this node.");
		
		return new PortObjectSpec[] { XlsFormatterStateSpec.getEmptySpec() };
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void reset() {
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void saveSettingsTo(final NodeSettingsWO settings) {
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadInternals(final File internDir,
			final ExecutionMonitor exec) throws IOException, CanceledExecutionException {
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void saveInternals(final File internDir,
			final ExecutionMonitor exec) throws IOException, CanceledExecutionException {
	}
}
