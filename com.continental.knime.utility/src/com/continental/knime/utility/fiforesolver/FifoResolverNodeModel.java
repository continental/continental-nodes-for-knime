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

package com.continental.knime.utility.fiforesolver;

import java.io.File;
import java.io.IOException;

import org.knime.core.data.DataColumnSpec;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataType;
import org.knime.core.data.DoubleValue;
import org.knime.core.data.def.StringCell;
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
import org.knime.core.node.defaultnodesettings.SettingsModelColumnName;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

public class FifoResolverNodeModel extends NodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(FifoResolverNodeModel.class);

	// Input Group Column Selection
	static final String CFGKEY_MODE = "QueueMode";
	static final String OPTION_MODE_FIFO = "FIFO";
	static final String OPTION_MODE_LIFO = "LIFO";
	static final String DEFAULT_MODE = OPTION_MODE_FIFO;
	final SettingsModelString m_mode =
			new SettingsModelString(CFGKEY_MODE, DEFAULT_MODE);
	
	// Input Group Column Selection
	static final String CFGKEY_COLUM_NAME_GROUP = "GroupCol";
	static final String DEFAULT_COLUMN_NAME_GROUP = "group_not_configured_but_model_default";
	final SettingsModelColumnName m_inputColumnNameGroup =
			new SettingsModelColumnName(CFGKEY_COLUM_NAME_GROUP, DEFAULT_COLUMN_NAME_GROUP);

	// Input Qty Column Selection
	static final String CFGKEY_COLUMN_NAME_QTY = "QtyCol";
	static final String DEFAULT_COLUMN_NAME_QTY = "quantity_not_configured_but_model_default";
	final SettingsModelColumnName m_inputColumnNameQty =
			new SettingsModelColumnName(CFGKEY_COLUMN_NAME_QTY, DEFAULT_COLUMN_NAME_QTY);

	// Error resolver boolean
	static final String CFGKEY_FAIL_AT_INCONSISTENCY = "FailAtInconsistency";
	static final boolean DEFAULT_FAIL_AT_INCONSISTENCY = false;
	final SettingsModelBoolean m_failAtInconsistency =
			new SettingsModelBoolean(CFGKEY_FAIL_AT_INCONSISTENCY, DEFAULT_FAIL_AT_INCONSISTENCY);   


	/**
	 * Constructor for the FifoResolverNodeModel.
	 */
	protected FifoResolverNodeModel() {

		// incoming port: BufferedDataTable with two columns
		// outgoing port: BufferedDataTable with group, qty, and rowid 
		super(1, 1);
	}

	int IN_PORT=0;

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected BufferedDataTable[] execute(final BufferedDataTable[] inData,
			final ExecutionContext exec) throws Exception {    		
		return FifoResolverNodeLogic.execute(
				inData,  // inData
				m_mode.getStringValue().equals(OPTION_MODE_FIFO),  // isModeFifo
				m_inputColumnNameGroup.getStringValue(),  // inputColumnNameGroup
				m_inputColumnNameQty.getStringValue(),  // inputColumnNameQty
				m_failAtInconsistency.getBooleanValue(),  // failAtInconsistency
				exec, IN_PORT, logger);
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
	protected DataTableSpec[] configure(final DataTableSpec[] inSpecs)
			throws InvalidSettingsException {

		// Check if input table has at least a String and a Numeric column and remember the first names
		String nameOfFirstStringColumn = null;
		String nameOfFirstNumericColumn = null;
		for (int i = 0; i < inSpecs[IN_PORT].getNumColumns(); i++) {
			DataColumnSpec columnSpec = inSpecs[IN_PORT].getColumnSpec(i);
			if (nameOfFirstStringColumn == null && columnSpec.getType().getCellClass().equals(StringCell.class))
				nameOfFirstStringColumn = columnSpec.getName();
			else if (nameOfFirstNumericColumn == null && columnSpec.getType().isCompatible(DoubleValue.class))
				nameOfFirstNumericColumn = columnSpec.getName();
		}

		if (nameOfFirstStringColumn == null || nameOfFirstNumericColumn == null)
			throw new InvalidSettingsException("Input table must contain at least a String and numeric column each.");
		
		if (m_inputColumnNameGroup.getStringValue().equals(DEFAULT_COLUMN_NAME_GROUP))
			m_inputColumnNameGroup.setStringValue(nameOfFirstStringColumn);
		if (m_inputColumnNameQty.getStringValue().equals(DEFAULT_COLUMN_NAME_QTY))
			m_inputColumnNameQty.setStringValue(nameOfFirstNumericColumn);
		
		String[] outNames = { m_inputColumnNameGroup.getStringValue(), "RowId_IN", "RowId_OUT", m_inputColumnNameQty.getStringValue() };
		DataType[] outTypes = { StringCell.TYPE, StringCell.TYPE, StringCell.TYPE, inSpecs[IN_PORT].getColumnSpec(nameOfFirstNumericColumn).getType() };
		DataColumnSpec[] colSpec= DataTableSpec.createColumnSpecs(outNames, outTypes);
		DataTableSpec outSpec = new DataTableSpec(colSpec);
		return new DataTableSpec[] { outSpec };
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void saveSettingsTo(final NodeSettingsWO settings) {

		m_mode.saveSettingsTo(settings);
		m_inputColumnNameGroup.saveSettingsTo(settings);
		m_inputColumnNameQty.saveSettingsTo(settings);
		m_failAtInconsistency.saveSettingsTo(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_mode.loadSettingsFrom(settings);
		m_inputColumnNameGroup.loadSettingsFrom(settings);
		m_inputColumnNameQty.loadSettingsFrom(settings);
		m_failAtInconsistency.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_mode.validateSettings(settings);
		m_inputColumnNameGroup.validateSettings(settings);
		m_inputColumnNameQty.validateSettings(settings);
		m_failAtInconsistency.validateSettings(settings);

		if (m_inputColumnNameGroup.getStringValue().equals(m_inputColumnNameQty.getStringValue()))
			throw new InvalidSettingsException("Two different input columns need to be selected.");
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
