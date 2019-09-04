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

package com.continental.knime.xlsformatter.xlscontroltablefromcellrange;

import java.io.File;
import java.io.IOException;

import org.apache.poi.ss.util.CellRangeAddress;
import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.MissingCell;
import org.knime.core.data.RowKey;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.BufferedDataContainer;
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

import com.continental.knime.xlsformatter.commons.AddressingTools;
import com.continental.knime.xlsformatter.commons.Commons.Modes;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableCreateTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;

public class XlsControlTableFromCellRangeNodeModel extends NodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsControlTableFromCellRangeNodeModel.class);


	static final String CFGKEY_CELLRANGE = "CellRange";
	static final String DEFAULT_CELLRANGE = "A1:B2";
	final SettingsModelString m_cellRange =
			new SettingsModelString(CFGKEY_CELLRANGE, DEFAULT_CELLRANGE);

	static final String CFGKEY_TAG = "Tag";
	static final String DEFAULT_TAG = "header";
	final SettingsModelString m_tag =
			new SettingsModelString(CFGKEY_TAG, DEFAULT_TAG);

	static final String CFGKEY_MODE = "Mode";
	static final String DEFAULT_MODE = Modes.APPEND.toString();
	final SettingsModelString m_mode =
			new SettingsModelString(CFGKEY_MODE, DEFAULT_MODE);

	/**
	 * Constructor for the node model.
	 */
	protected XlsControlTableFromCellRangeNodeModel() {
		super(
				new PortType[] { BufferedDataTable.TYPE_OPTIONAL },
				new PortType[] { BufferedDataTable.TYPE });
	}


	/**
	 * {@inheritDoc}
	 */
	protected BufferedDataTable[] execute(BufferedDataTable[] inData,
			final ExecutionContext exec) throws Exception {

		CellRangeAddress range;
		try {
			range = AddressingTools.parseRange(m_cellRange.getStringValue().trim());
		}
		catch (Exception e) {
			throw new Exception("Invalid cell range " + m_cellRange.getStringValue().trim() + ". " + e.getMessage(), e);
		}

		if (!XlsFormatterControlTableValidator.isSheetSizeWithinXlsSpec(range.getLastColumn() + 1, range.getLastRow() + 1))
			throw new Exception("Cell range exceeds XLS limits: " + m_cellRange.getStringValue().trim());


		// generate new control table:
		BufferedDataContainer newDataTableBuffer = XlsFormatterControlTableCreateTools.getNewBufferedDataContainer(range.getLastColumn() + 1, exec, logger);
		RowKey[] rowKeyArr = XlsFormatterControlTableCreateTools.getRowkeyArray(range.getLastRow() + 1, exec, logger);

		for (int r = 0; r <= range.getLastRow(); r++) {
			DataCell[] cells = new DataCell[range.getLastColumn() + 1];
			for (int c = 0; c <= range.getLastColumn(); c++)
				cells[c] = range.containsColumn(c) && range.containsRow(r) ? new StringCell(m_tag.getStringValue().trim()) : new MissingCell("");
				DataRow rowOut = new DefaultRow(rowKeyArr[r], cells);
				newDataTableBuffer.addRowToTable(rowOut);
		}
		newDataTableBuffer.close();
		BufferedDataTable newDataTable = newDataTableBuffer.getTable();

		// in case of an existing table coming in, merge:
		if (inData[0] != null)
			newDataTable = XlsFormatterControlTableCreateTools.merge(
					inData[0], newDataTable,
					m_mode.getStringValue().equals(Modes.OVERWRITE.toString()), exec, logger);

		return new BufferedDataTable[] { newDataTable };
	}

	@Override
	protected DataTableSpec[] configure(final DataTableSpec[] inSpecs)
			throws InvalidSettingsException {

		if (inSpecs[0] != null && !XlsFormatterControlTableValidator.isControlTableSpec(inSpecs[0], logger))
			throw new InvalidSettingsException("The input table specification is not that of a valid XLS Formatting control table. See log for details.");

		CellRangeAddress range;
		try {
			range = AddressingTools.parseRange(m_cellRange.getStringValue().trim());
		}
		catch (Exception e) {
			throw new InvalidSettingsException("Invalid cell range " + m_cellRange.getStringValue().trim() + ". " + e.getMessage(), e);
		}

		if (inSpecs[0] == null)
			return new DataTableSpec[] { XlsFormatterControlTableCreateTools.createDataTableSpec(range.getLastColumn() + 1) };

		return new DataTableSpec[] { XlsFormatterControlTableCreateTools.createDataTableSpec(
				Math.max(inSpecs[0].getNumColumns(), range.getLastColumn() + 1)) };
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
		m_cellRange.saveSettingsTo(settings);
		m_tag.saveSettingsTo(settings);
		m_mode.saveSettingsTo(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_cellRange.loadSettingsFrom(settings);
		m_tag.loadSettingsFrom(settings);    
		m_mode.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_cellRange.validateSettings(settings);
		m_tag.validateSettings(settings);
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