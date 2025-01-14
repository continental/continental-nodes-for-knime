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

package com.continental.knime.utility.networkcomponentsplitter;

import java.io.File;
import java.io.IOException;

import org.knime.core.data.DataColumnSpec;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataType;
import org.knime.core.data.def.IntCell;
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

public class NetworkComponentSplitterNodeModel extends NodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(NetworkComponentSplitterNodeModel.class);

	// Input Node1 Column Selection
	static final String CFGKEY_COLUM_NAME1 = "InputCol1";
	static final String DEFAULT_COLUMN_NAME1 = "Node1_not_configured_but_model_default";
	final SettingsModelColumnName m_inputColumnName1 =
			new SettingsModelColumnName(CFGKEY_COLUM_NAME1, DEFAULT_COLUMN_NAME1);

	// Input Node2 Column Selection
	static final String CFGKEY_COLUMN_NAME2 = "InputCol2";
	static final String DEFAULT_COLUMN_NAME2 = "Node2_not_configured_but_model_default";
	final SettingsModelColumnName m_inputColumnName2 =
			new SettingsModelColumnName(CFGKEY_COLUMN_NAME2, DEFAULT_COLUMN_NAME2);

	// Output Node ColumnName
	static final String CFGKEY_OUTPUT_COLUMN_NAME_NODE = "OutputColNode";
	static final String DEFAULT_OUTPUT_COLUMN_NAME_NODE = "Node";
	static final SettingsModelString m_outputColumnNameNode =
			new SettingsModelString(CFGKEY_OUTPUT_COLUMN_NAME_NODE, DEFAULT_OUTPUT_COLUMN_NAME_NODE);

	// Output Cluster ColumnName
	static final String CFGKEY_OUTPUT_COLUMN_NAME_CLUSTER = "OutputColCluster";
	static final String DEFAULT_OUTPUT_COLUMN_NAME_CLUSTER = "Cluster";
	final SettingsModelString m_outputColumnNameCluster =
			new SettingsModelString(CFGKEY_OUTPUT_COLUMN_NAME_CLUSTER, DEFAULT_OUTPUT_COLUMN_NAME_CLUSTER);

	// Missing Value Handling TickMark
	static final String CFGKEY_MISSING = "MissingAsOwnNode";
	static final boolean DEFAULT_MISSING = true;
	final SettingsModelBoolean m_missingValueAllowedAsOwnNode =
			new SettingsModelBoolean(CFGKEY_MISSING, DEFAULT_MISSING);   


	/**
	 * Constructor for the NetworkComponentSplitterNodeModel.
	 */
	protected NetworkComponentSplitterNodeModel() {

		// incoming port: BufferedDataTable with two columns of connected Nodes
		// outgoing port: BufferedDataTable with Nodes as String and ClusterID as int 
		super(1, 1);
	}

	int IN_PORT=0;

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected BufferedDataTable[] execute(final BufferedDataTable[] inData,
			final ExecutionContext exec) throws Exception {    		
		return NetworkComponentSplitterNodeLogic.execute(inData, m_inputColumnName1.getStringValue(), m_inputColumnName2.getStringValue(),
				m_missingValueAllowedAsOwnNode.getBooleanValue(), m_outputColumnNameNode.getStringValue().trim(),
				m_outputColumnNameCluster.getStringValue().trim(), exec, IN_PORT, logger);
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

		// Check if input table has at least two String columns and remember the first two String columns' names
		int numberStringColumns = 0;
		String[] inNames = new String[2];
		for (int i = 0; i < inSpecs[IN_PORT].getNumColumns(); i++) {
			DataColumnSpec columnSpec = inSpecs[IN_PORT].getColumnSpec(i);
			if (columnSpec.getType().getCellClass().equals(StringCell.class)) {
				inNames[numberStringColumns] = columnSpec.getName();
				numberStringColumns++;
			}
			if (numberStringColumns == 2)
				break;
		}

		if (numberStringColumns < 2)
			throw new InvalidSettingsException("Input table must contain at least two String columns.");

		if (m_inputColumnName1.getStringValue().equals(DEFAULT_COLUMN_NAME1))
			m_inputColumnName1.setStringValue(inNames[0]);
		if (m_inputColumnName2.getStringValue().equals(DEFAULT_COLUMN_NAME2))
			m_inputColumnName2.setStringValue(inNames[1]);
		
		String[] outNames = { m_outputColumnNameNode.getStringValue(), m_outputColumnNameCluster.getStringValue() };
		DataType[] outTypes = { StringCell.TYPE, IntCell.TYPE };
		DataColumnSpec[] colSpec= DataTableSpec.createColumnSpecs(outNames, outTypes);
		DataTableSpec outSpec = new DataTableSpec(colSpec);
		return new DataTableSpec[] { outSpec };
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void saveSettingsTo(final NodeSettingsWO settings) {

		m_inputColumnName1.saveSettingsTo(settings);
		m_inputColumnName2.saveSettingsTo(settings);
		m_missingValueAllowedAsOwnNode.saveSettingsTo(settings);
		m_outputColumnNameNode.saveSettingsTo(settings);
		m_outputColumnNameCluster.saveSettingsTo(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_inputColumnName1.loadSettingsFrom(settings);
		m_inputColumnName2.loadSettingsFrom(settings);
		m_missingValueAllowedAsOwnNode.loadSettingsFrom(settings);
		m_outputColumnNameNode.loadSettingsFrom(settings);
		m_outputColumnNameCluster.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_inputColumnName1.validateSettings(settings);
		m_inputColumnName2.validateSettings(settings);
		m_missingValueAllowedAsOwnNode.validateSettings(settings);
		m_outputColumnNameNode.validateSettings(settings);
		m_outputColumnNameCluster.validateSettings(settings);

		if (m_inputColumnName1.getStringValue().equals(m_inputColumnName2.getStringValue()))
			throw new InvalidSettingsException("Two different input columns need to be selected.");
		
		if (m_outputColumnNameNode.getStringValue().trim().equals(""))
			throw new InvalidSettingsException("The name of the first output column cannot be empty.");
		
		if (m_outputColumnNameCluster.getStringValue().trim().equals(""))
			throw new InvalidSettingsException("The name of the second output column cannot be empty.");
		
		if (m_outputColumnNameNode.getStringValue().trim().equals(m_outputColumnNameCluster.getStringValue().trim()))
			throw new InvalidSettingsException("Two different output columns names need to be defined.");
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
