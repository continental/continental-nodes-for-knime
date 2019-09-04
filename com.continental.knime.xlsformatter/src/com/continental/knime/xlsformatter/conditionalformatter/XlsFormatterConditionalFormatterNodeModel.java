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
import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
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
	static final String CFGKEY_TAG = "Tag";
	static final String DEFAULT_TAG = "header";
	final SettingsModelString m_tag =
			new SettingsModelString(CFGKEY_TAG, DEFAULT_TAG);

	static final String CFGKEY_MIDSCALEPOINT_ACTIVE = "MidScalePointOn";
	static final boolean DEFAULT_MIDSCALEPOINT_ACTIVE = true;
	final SettingsModelBoolean m_midScalePointActive =
			new SettingsModelBoolean(CFGKEY_MIDSCALEPOINT_ACTIVE, DEFAULT_MIDSCALEPOINT_ACTIVE);

	static final String CFGKEY_MIN_THRESHOLD = "MinValue";
	static final double DEFAULT_MIN_THRESHOLD = 0.0;
	final SettingsModelDouble m_minThreshold =
			new SettingsModelDouble(CFGKEY_MIN_THRESHOLD, DEFAULT_MIN_THRESHOLD);   

	static final String CFGKEY_MIN_COLOR = "MinColor";
	static final Color DEFAULT_MIN_COLOR = new Color(0, 255, 0);
	final SettingsModelColor m_minColor =
			new SettingsModelColor(CFGKEY_MIN_COLOR,DEFAULT_MIN_COLOR);

	static final String CFGKEY_MID_THRESHOLD = "MidValue";
	static final double DEFAULT_MID_THRESHOLD = 0.5;
	final SettingsModelDouble m_midThreshold =
			new SettingsModelDouble(CFGKEY_MID_THRESHOLD, DEFAULT_MID_THRESHOLD);   

	static final String CFGKEY_MID_COLOR = "MidColor";
	static final Color DEFAULT_MID_COLOR = new Color(255, 255, 0);
	final SettingsModelColor m_midColor =
			new SettingsModelColor(CFGKEY_MID_COLOR, DEFAULT_MID_COLOR);

	static final String CFGKEY_MAX_THRESHOLD = "MaxValue";
	static final double DEFAULT_MAX_THRESHOLD = 1.0;
	final SettingsModelDouble m_maxThreshold =
			new SettingsModelDouble(CFGKEY_MAX_THRESHOLD, DEFAULT_MAX_THRESHOLD);   

	static final String CFGKEY_MAX_COLOR = "MaxColor";
	static final Color DEFAULT_MAX_COLOR = new Color(255, 0, 0);
	final SettingsModelColor m_maxColor =
			new SettingsModelColor(CFGKEY_MAX_COLOR, DEFAULT_MAX_COLOR);

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
		XlsFormatterState.SheetState xlsfs = xlsf.getCurrentSheetStateForModification();
		WarningMessageContainer warningMessageContainer = new WarningMessageContainer();

		List<CellAddress> targetCells =
				XlsFormatterControlTableAnalysisTools.getCellsMatchingTag((BufferedDataTable)inObjects[0], m_tag.getStringValue().trim(), exec, logger);
		warnOnNoMatchingTags(targetCells, m_tag.getStringValue().trim(), warningMessageContainer);

		// check for a partly overlap of these target cells with a previously merged range and warn
		String mergeOverlapRanges = XlsFormatterControlTableAnalysisTools.getOverlappingRanges(targetCells, xlsfs.mergeRanges, exec, logger);
		if (mergeOverlapRanges != null)
			warningMessageContainer.addMessage("Modification on parts of previously merged range(s) (" + mergeOverlapRanges + ") will have no effect.");
		
		// translate UI conditional format definition to internal representation:
		ConditionalFormattingSet condFormatSet = new ConditionalFormattingSet();
		condFormatSet.backgroundScaleFixpoints.add(Pair.of(m_minThreshold.getDoubleValue(), m_minColor.getColorValue()));
		if (m_midScalePointActive.getBooleanValue())
			condFormatSet.backgroundScaleFixpoints.add(Pair.of(m_midThreshold.getDoubleValue(), m_midColor.getColorValue()));
		condFormatSet.backgroundScaleFixpoints.add(Pair.of(m_maxThreshold.getDoubleValue(), m_maxColor.getColorValue()));
		
		
		for (CellAddress cell : targetCells) {
			XlsFormatterState.CellState cellState = AddressingTools.safelyGetCellInMap(xlsfs.cells, cell);
			cellState.conditionalFormat = condFormatSet; // note that we share the ConditionalFormattingSet object here across cells!
		}
		
		if (warningMessageContainer.hasMessage())
			setWarningMessage(warningMessageContainer.getMessage());
		
		return new PortObject[] { xlsf };
	}

	@Override
	protected PortObjectSpec[] configure(final PortObjectSpec[] inSpecs)
			throws InvalidSettingsException {   	

		if (!XlsFormatterControlTableValidator.isControlTableSpec((DataTableSpec)inSpecs[0], logger))
			throw new InvalidSettingsException("The configured input table header is not that of a valid XLS Formatting control table. See log for details.");

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

		m_tag.saveSettingsTo(settings);
		m_midScalePointActive.saveSettingsTo(settings);
		m_minThreshold.saveSettingsTo(settings);
		m_minColor.saveSettingsTo(settings);
		m_midThreshold.saveSettingsTo(settings);
		m_midColor.saveSettingsTo(settings);
		m_maxThreshold.saveSettingsTo(settings);
		m_maxColor.saveSettingsTo(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_tag.loadSettingsFrom(settings);
		m_midScalePointActive.loadSettingsFrom(settings);
		m_minThreshold.loadSettingsFrom(settings);
		m_minColor.loadSettingsFrom(settings);
		m_midThreshold.loadSettingsFrom(settings);
		m_midColor.loadSettingsFrom(settings);
		m_maxThreshold.loadSettingsFrom(settings);
		m_maxColor.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_tag.validateSettings(settings);
		m_midScalePointActive.validateSettings(settings);
		m_minThreshold.validateSettings(settings);
		m_minColor.validateSettings(settings);
		m_midThreshold.validateSettings(settings);
		m_midColor.validateSettings(settings);
		m_maxThreshold.validateSettings(settings);
		m_maxColor.validateSettings(settings);
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
