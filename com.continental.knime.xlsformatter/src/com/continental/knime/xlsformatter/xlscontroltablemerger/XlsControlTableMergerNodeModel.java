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

package com.continental.knime.xlsformatter.xlscontroltablemerger;

import java.io.File;
import java.io.IOException;
import java.util.Arrays;

import org.knime.core.data.DataTableSpec;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.commons.Commons.Modes;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableCreateTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;

public class XlsControlTableMergerNodeModel extends NodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsControlTableMergerNodeModel.class);

	static final String CFGKEY_MODE = "Mode";
	static final String DEFAULT_MODE = Modes.APPEND.toString();
	final SettingsModelString m_mode =
			new SettingsModelString(CFGKEY_MODE, DEFAULT_MODE);

	/**
	 * Constructor for the node model.
	 */
	public XlsControlTableMergerNodeModel() {
		super(2, 1);
	}
	XlsControlTableMergerNodeModel(final int nrIns) {
    super(getInPortTypes(nrIns), new PortType[] {BufferedDataTable.TYPE});
  }
	
	XlsControlTableMergerNodeModel(final PortsConfiguration portsConfiguration) {
    super(portsConfiguration.getInputPorts(), portsConfiguration.getOutputPorts());
	}

	private static final PortType[] getInPortTypes(final int nrIns) {
    if (nrIns < 2) {
    	throw new IllegalArgumentException("Invalid number of input tables (" + nrIns + "). Merge operation requires at least 2 inputs.");
    }
    PortType[] result = new PortType[nrIns];
    Arrays.fill(result, BufferedDataTable.TYPE_OPTIONAL);
    result[0] = BufferedDataTable.TYPE;
    result[1] = BufferedDataTable.TYPE;
    return result;
}

	/**
	 * {@inheritDoc}
	 */
	protected BufferedDataTable[] execute(BufferedDataTable[] inData, 
			final ExecutionContext exec) throws Exception {

		for (int i = 0; i < inData.length; i++)
			if (!XlsFormatterControlTableValidator.isControlTable(inData[i], exec, logger))
				throw new Exception("Input table " + (i+1) + " is not a XLS Formatter Control Table. See log for details.");

		BufferedDataTable currentTable = inData[0];
		for (int i = 1; i < inData.length; i++)
			currentTable = XlsFormatterControlTableCreateTools.merge(currentTable, inData[i],
					Modes.getFromString(m_mode.getStringValue()) == Modes.OVERWRITE, exec, logger);
		
		return new BufferedDataTable[] { currentTable };
	}


	@Override
	protected DataTableSpec[] configure(final DataTableSpec[] inSpecs)
			throws InvalidSettingsException {

		int width = -1;
		for (int i = 0; i < inSpecs.length; i++) {
			if (!XlsFormatterControlTableValidator.isControlTableSpec(inSpecs[i], logger))
				throw new InvalidSettingsException("Input table header " + (i+1) + " is not that of a valid XLS Formatting control table. See log for details.");
			
			if (inSpecs[i].getNumColumns() > width)
				width = inSpecs[i].getNumColumns();
		}

		return new DataTableSpec[] { XlsFormatterControlTableCreateTools.createDataTableSpec(width) };
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

		m_mode.saveSettingsTo(settings);        
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_mode.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_mode.validateSettings(settings);
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
