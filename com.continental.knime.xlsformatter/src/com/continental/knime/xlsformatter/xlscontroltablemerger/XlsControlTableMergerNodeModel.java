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
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.commons.Commons.Modes;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableCreateTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;

public class XlsControlTableMergerNodeModel extends NodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsControlTableMergerNodeModel.class);

	static final String CFGKEY_MODE_STRING = "Mode";
	static final String DEFAULT_MODE_STRING = Modes.APPEND.toString();
	final SettingsModelString m_modeString =
			new SettingsModelString(CFGKEY_MODE_STRING, DEFAULT_MODE_STRING);

	/**
	 * Constructor for the node model.
	 */
	protected XlsControlTableMergerNodeModel() {

		super(
				new PortType[] { BufferedDataTable.TYPE, BufferedDataTable.TYPE },
				new PortType[] { BufferedDataTable.TYPE });
	}

	/**
	 * {@inheritDoc}
	 */
	protected BufferedDataTable[] execute(BufferedDataTable[] inData, 
			final ExecutionContext exec) throws Exception {

		if (!XlsFormatterControlTableValidator.isControlTable(inData[0], exec, logger))
			throw new Exception("Top table is not a XLS Formatter Control Table.");
		if (!XlsFormatterControlTableValidator.isControlTable(inData[1], exec, logger))
			throw new Exception("Bottom table is not a XLS Formatter Control Table.");

		return new BufferedDataTable[] { XlsFormatterControlTableCreateTools.merge(inData[0], inData[1],
				Modes.getFromString(m_modeString.getStringValue()) == Modes.OVERWRITE, exec, logger) };
	}


	@Override
	protected DataTableSpec[] configure(final DataTableSpec[] inSpecs)
			throws InvalidSettingsException {

		if (!XlsFormatterControlTableValidator.isControlTableSpec(inSpecs[0], logger))
			throw new InvalidSettingsException("The top input table header is not that of a valid XLS Formatting control table. See log for details.");
		if (!XlsFormatterControlTableValidator.isControlTableSpec(inSpecs[1], logger))
			throw new InvalidSettingsException("The bottom input table header is not that of a valid XLS Formatting control table. See log for details.");

		int topWidth = inSpecs[0].getNumColumns();
		int bottomWidth = inSpecs[1].getNumColumns();
		int width = Math.max(topWidth, bottomWidth);

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

		m_modeString.saveSettingsTo(settings);        
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_modeString.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_modeString.validateSettings(settings);
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
