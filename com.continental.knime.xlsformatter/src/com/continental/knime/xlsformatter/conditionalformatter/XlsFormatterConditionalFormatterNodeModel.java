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

package com.continental.knime.xlsformatter.conditionalformatter;

import java.awt.Color;
import java.io.File;
import java.io.IOException;
import java.util.List;

import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.util.CellAddress;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelColor;
import org.knime.core.node.defaultnodesettings.SettingsModelDouble;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.commons.AddressingTools;
import com.continental.knime.xlsformatter.commons.TagBasedXlsCellFormatterNodeModel;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableAnalysisTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.ConditionalFormattingSet;
import com.continental.knime.xlsformatter.porttype.XlsFormatterStateSpec;

public class XlsFormatterConditionalFormatterNodeModel extends TagBasedXlsCellFormatterNodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsFormatterConditionalFormatterNodeModel.class);

	//Input tag string
	static final String CFGKEY_TAGSTRING = "Tag";
	static final String DEFAULT_TAGSTRING = "header";
	final SettingsModelString m_tagstring =
			new SettingsModelString(CFGKEY_TAGSTRING,DEFAULT_TAGSTRING);

	static final String CFGKEY_MIDSCALEPOINT_ON = "MidScalePointOn";
	static final boolean DEFAULT_MIDSCALEPOINT_ON = true;
	final SettingsModelBoolean m_midscalepointon =
			new SettingsModelBoolean(CFGKEY_MIDSCALEPOINT_ON,DEFAULT_MIDSCALEPOINT_ON);

	static final String CFGKEY_MIN_VALUE = "MinValue";
	static final double DEFAULT_MIN_VALUE = 0.0;
	final SettingsModelDouble m_minvalue =
			new SettingsModelDouble(CFGKEY_MIN_VALUE,DEFAULT_MIN_VALUE);   

	static final String CFGKEY_MIN_COLOR = "MinColor";
	static final Color DEFAULT_MIN_COLOR = new Color(0,255,0);
	final SettingsModelColor m_mincolor =
			new SettingsModelColor(CFGKEY_MIN_COLOR,DEFAULT_MIN_COLOR);

	static final String CFGKEY_MID_VALUE = "MidValue";
	static final double DEFAULT_MID_VALUE = 0.5;
	final SettingsModelDouble m_midvalue =
			new SettingsModelDouble(CFGKEY_MID_VALUE,DEFAULT_MID_VALUE);   

	static final String CFGKEY_MID_COLOR = "MidColor";
	static final Color DEFAULT_MID_COLOR = new Color(255,255,0);
	final SettingsModelColor m_midcolor =
			new SettingsModelColor(CFGKEY_MID_COLOR,DEFAULT_MID_COLOR);

	static final String CFGKEY_MAX_VALUE = "MaxValue";
	static final double DEFAULT_MAX_VALUE = 1.0;
	final SettingsModelDouble m_maxvalue =
			new SettingsModelDouble(CFGKEY_MAX_VALUE,DEFAULT_MAX_VALUE);   

	static final String CFGKEY_MAX_COLOR = "MaxColor";
	static final Color DEFAULT_MAX_COLOR = new Color(255,0,0);
	final SettingsModelColor m_maxcolor =
			new SettingsModelColor(CFGKEY_MAX_COLOR,DEFAULT_MAX_COLOR);

	/**
	 * Constructor for the node model.
	 */
	protected XlsFormatterConditionalFormatterNodeModel() {

		super(
				new PortType[] { BufferedDataTable.TYPE, XlsFormatterState.TYPE_OPTIONAL },
				new PortType[] { XlsFormatterState.TYPE });
	}

	/**
	 * {@inheritDoc}
	 */
	protected PortObject[] execute(PortObject[] inObjects,
			final ExecutionContext exec) throws Exception {

		if (!XlsFormatterControlTableValidator.isControlTable((BufferedDataTable)inObjects[0], exec, logger))
			throw new IllegalArgumentException("The provided input table is not a valid XLS control table. See log for details.");

		XlsFormatterState xlsf = XlsFormatterState.getDeepClone(inObjects[1]);

		List<CellAddress> targetCells =
				XlsFormatterControlTableAnalysisTools.getCellsMatchingTag((BufferedDataTable)inObjects[0], m_tagstring.getStringValue().trim(), exec, logger);
		warnOnNoMatchingTags(targetCells, m_tagstring.getStringValue().trim());

		// translate UI conditional format definition to internal representation:
		ConditionalFormattingSet condFormatSet = new ConditionalFormattingSet();
		condFormatSet.backgroundScaleFixpoints.add(Pair.of(m_minvalue.getDoubleValue(), m_mincolor.getColorValue()));
		if (m_midscalepointon.getBooleanValue())
			condFormatSet.backgroundScaleFixpoints.add(Pair.of(m_midvalue.getDoubleValue(), m_midcolor.getColorValue()));
		condFormatSet.backgroundScaleFixpoints.add(Pair.of(m_maxvalue.getDoubleValue(), m_maxcolor.getColorValue()));
		
		
		for (CellAddress cell : targetCells) {
			XlsFormatterState.CellState cellState = AddressingTools.safelyGetCellInMap(xlsf.cells, cell);
			cellState.conditionalFormat = condFormatSet; // note that we share the ConditionalFormattingSet object here across cells!
		}
		
		return new PortObject[] { xlsf };
	}

	@Override
	protected PortObjectSpec[] configure(final PortObjectSpec[] inSpecs)
			throws InvalidSettingsException {   	

		if (!XlsFormatterControlTableValidator.isControlTableSpec((DataTableSpec)inSpecs[0], logger))
			throw new InvalidSettingsException("The configured input table header is not that of a valid XLS Formatting control table. See log for details.");

		if (inSpecs[1] != null && ((XlsFormatterStateSpec)inSpecs[1]).getContainsMergeInstruction() == true)
			throw new InvalidSettingsException("No futher XLS Formatting nodes allowed after Cell Merger.");

		return new PortObjectSpec[] { inSpecs[1] == null ? XlsFormatterStateSpec.getEmptySpec() : ((XlsFormatterStateSpec)inSpecs[1]).getCopy() };
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

		m_tagstring.saveSettingsTo(settings);
		m_midscalepointon.saveSettingsTo(settings);
		m_minvalue.saveSettingsTo(settings);
		m_mincolor.saveSettingsTo(settings);
		m_midvalue.saveSettingsTo(settings);
		m_midcolor.saveSettingsTo(settings);
		m_maxvalue.saveSettingsTo(settings);
		m_maxcolor.saveSettingsTo(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_tagstring.loadSettingsFrom(settings);
		m_midscalepointon.loadSettingsFrom(settings);
		m_minvalue.loadSettingsFrom(settings);
		m_mincolor.loadSettingsFrom(settings);
		m_midvalue.loadSettingsFrom(settings);
		m_midcolor.loadSettingsFrom(settings);
		m_maxvalue.loadSettingsFrom(settings);
		m_maxcolor.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_tagstring.validateSettings(settings);
		m_midscalepointon.validateSettings(settings);
		m_minvalue.validateSettings(settings);
		m_mincolor.validateSettings(settings);
		m_midvalue.validateSettings(settings);
		m_midcolor.validateSettings(settings);
		m_maxvalue.validateSettings(settings);
		m_maxcolor.validateSettings(settings);
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
