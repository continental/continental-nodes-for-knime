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

package com.continental.knime.xlsformatter.sheetselector;

import java.io.File;
import java.io.IOException;

import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterStateSpec;

public class XlsSheetSelectorNodeModel extends NodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsSheetSelectorNodeModel.class);

	static final String CFGKEY_SHEET_NAME = "SheetName";
	static final String DEFAULT_SHEET_NAME = "Table 1";
	final SettingsModelString m_sheetName =
			new SettingsModelString(CFGKEY_SHEET_NAME, DEFAULT_SHEET_NAME);

	static final String CFGKEY_OPTION_NAMED = "OptionNamed";
	static final boolean DEFAULT_OPTION_NAMED = true;
	/**
	 * True means that the provided sheet name is used, false means that the default sheet 0 is addressed.
	 */
	final SettingsModelBoolean m_optionNamed =
			new SettingsModelBoolean(CFGKEY_OPTION_NAMED, DEFAULT_OPTION_NAMED);


	/**
	 * Constructor for the node model.
	 */
	protected XlsSheetSelectorNodeModel() {
		super(
				new PortType[] { },
				new PortType[] { XlsFormatterState.TYPE });
	}

	/**
	 * {@inheritDoc}
	 */
	protected PortObject[] execute(PortObject[] inObjects, final ExecutionContext exec) throws Exception { 

		XlsFormatterState xlsf = new XlsFormatterState();
		
		if (m_sheetName.getStringValue().trim().equals(""))
			throw new IllegalArgumentException("An empty sheet name is not allowed");
		
		if (m_optionNamed.getBooleanValue()) // in case of false, the default logic in the PortState implementation holds
			xlsf.setCurrentSheetForModification(m_sheetName.getStringValue());
		
		logger.info("The sheet for further instructions has been changed to " +
				(m_optionNamed.getBooleanValue() ? "\"" + m_sheetName.getStringValue() + "\"" : "[default: sheet 0]"));
		
		return new PortObject[] { xlsf };
	}

	@Override
	protected PortObjectSpec[] configure(final PortObjectSpec[] inSpecs)
			throws InvalidSettingsException {   	

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

		m_sheetName.saveSettingsTo(settings);
		m_optionNamed.saveSettingsTo(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_sheetName.loadSettingsFrom(settings);
		m_optionNamed.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_sheetName.validateSettings(settings);		
		m_optionNamed.validateSettings(settings);
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
