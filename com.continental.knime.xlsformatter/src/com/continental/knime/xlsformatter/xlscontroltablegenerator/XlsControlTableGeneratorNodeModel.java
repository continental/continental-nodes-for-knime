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

package com.continental.knime.xlsformatter.xlscontroltablegenerator;

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
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.commons.WarningMessageContainer;

public class XlsControlTableGeneratorNodeModel extends NodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsControlTableGeneratorNodeModel.class);

	static final String CFGKEY_ROWSHIFT = "RowShift";
	static final boolean DEFAULT_ROWSHIFT = false;
	final SettingsModelBoolean m_rowShift =
			new SettingsModelBoolean(CFGKEY_ROWSHIFT, DEFAULT_ROWSHIFT);

	static final String CFGKEY_UNPIVOT = "Unpivot";
	static final boolean DEFAULT_UNPIVOT = false;
	final SettingsModelBoolean m_unpivot =
			new SettingsModelBoolean(CFGKEY_UNPIVOT, DEFAULT_UNPIVOT);

	static final String CFGKEY_EXTENDEDUNPIVOTCOLUMNS = "ExtendedHeaderColumns";
	static final boolean DEFAULT_EXTENDEDUNPIVOTCOLUMNS = true;
	final SettingsModelBoolean m_extendedColumnsAtUnpivot =
			new SettingsModelBoolean(CFGKEY_EXTENDEDUNPIVOTCOLUMNS, DEFAULT_EXTENDEDUNPIVOTCOLUMNS);

	/**
	 * Constructor for the node model.
	 */
	protected XlsControlTableGeneratorNodeModel() {

		super(
				new PortType[] { BufferedDataTable.TYPE },
				new PortType[] { BufferedDataTable.TYPE });
	}

	/**
	 * {@inheritDoc}
	 */
	protected BufferedDataTable[] execute(BufferedDataTable[] inData,
			final ExecutionContext exec) throws Exception {
		
		if (m_unpivot.getBooleanValue())
			return XlsControlTableGeneratorFunctionUnpivot.unpivot(inData[0], m_rowShift.getBooleanValue(), exec, logger, m_extendedColumnsAtUnpivot.getBooleanValue());
		else {
			WarningMessageContainer warningMessageContainer = new WarningMessageContainer();
			BufferedDataTable[] ret = XlsControlTableGeneratorFunctionDerivePivoted.derivePivoted(
					inData[0], m_rowShift.getBooleanValue(), true, warningMessageContainer, exec, logger);
			if (warningMessageContainer.hasMessage())
				setWarningMessage(warningMessageContainer.getMessage());
			return ret;
		}
	}

	@Override
	protected DataTableSpec[] configure(final DataTableSpec[] inSpecs)
			throws InvalidSettingsException {

		if (m_unpivot.getBooleanValue())
			return new DataTableSpec[] { XlsControlTableGeneratorFunctionUnpivot.getUnpivotSpec(inSpecs[0], m_extendedColumnsAtUnpivot.getBooleanValue()) };

		return new DataTableSpec[] { XlsControlTableGeneratorFunctionDerivePivoted.getPivotedSpec(inSpecs[0]) };
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

		m_extendedColumnsAtUnpivot.setEnabled(m_unpivot.getBooleanValue());
		
		m_rowShift.saveSettingsTo(settings);
		m_unpivot.saveSettingsTo(settings);
		m_extendedColumnsAtUnpivot.saveSettingsTo(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_rowShift.loadSettingsFrom(settings);
		m_unpivot.loadSettingsFrom(settings);
		m_extendedColumnsAtUnpivot.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_rowShift.validateSettings(settings);
		m_unpivot.validateSettings(settings);
		m_extendedColumnsAtUnpivot.validateSettings(settings);
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
