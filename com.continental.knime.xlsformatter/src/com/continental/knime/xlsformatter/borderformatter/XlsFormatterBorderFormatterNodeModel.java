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

package com.continental.knime.xlsformatter.borderformatter;

import java.awt.Color;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

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
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.commons.TagBasedXlsCellFormatterNodeModel;
import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableAnalysisTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator.ControlTableType;
import com.continental.knime.xlsformatter.commons.XlsFormattingStateValidator;
import com.continental.knime.xlsformatter.commons.XlsFormattingStateValidator.ValidationModes;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterStateSpec;

public class XlsFormatterBorderFormatterNodeModel extends TagBasedXlsCellFormatterNodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsFormatterBorderFormatterNodeModel.class);

	//Input tag string
	static final String CFGKEY_TAGSTRING = "Tag";
	static final String DEFAULT_TAGSTRING = "header";
	final SettingsModelString m_tag =
			new SettingsModelString(CFGKEY_TAGSTRING, DEFAULT_TAGSTRING);

	static final String CFGKEY_ALL_TAGS = "AllTags";
	static final boolean DEFAULT_ALL_TAGS = false;
	final SettingsModelBoolean m_allTags =
			new SettingsModelBoolean(CFGKEY_ALL_TAGS, DEFAULT_ALL_TAGS);

	static final String CFGKEY_BORDER_STYLE = "BorderStyle";
	static final String DEFAULT_BORDER_STYLE = XlsFormatterState.BorderStyle.NORMAL.toString().toLowerCase();
	final SettingsModelString m_borderStyle =
			new SettingsModelString(CFGKEY_BORDER_STYLE, DEFAULT_BORDER_STYLE);

	static final String CFGKEY_BORDER_CHANGECOLOR = "ChangeBorderColor";
	static final boolean DEFAULT_BORDER_CHANGECOLOR = false;
	final SettingsModelBoolean m_borderChangeColor =
			new SettingsModelBoolean(CFGKEY_BORDER_CHANGECOLOR, DEFAULT_BORDER_CHANGECOLOR);

	static final String CFGKEY_BORDER_COLOR = "BorderColor";
	static final Color DEFAULT_BORDER_COLOR = new Color(0, 0, 0);
	final SettingsModelColor m_borderColor =
			new SettingsModelColor(CFGKEY_BORDER_COLOR, DEFAULT_BORDER_COLOR);

	static final String CFGKEY_BORDER_TOP = "BorderTop";
	static final boolean DEFAULT_BORDER_TOP = false;
	final SettingsModelBoolean m_borderTop =
			new SettingsModelBoolean(CFGKEY_BORDER_TOP, DEFAULT_BORDER_TOP);

	static final String CFGKEY_BORDER_BOTTOM = "BorderBottom";
	static final boolean DEFAULT_BORDER_BOTTOM = false;
	final SettingsModelBoolean m_borderBottom =
			new SettingsModelBoolean(CFGKEY_BORDER_BOTTOM, DEFAULT_BORDER_BOTTOM);

	static final String CFGKEY_BORDER_LEFT = "BorderLeft";
	static final boolean DEFAULT_BORDER_LEFT = false;
	final SettingsModelBoolean m_borderLeft =
			new SettingsModelBoolean(CFGKEY_BORDER_LEFT, DEFAULT_BORDER_LEFT);

	static final String CFGKEY_BORDER_RIGHT = "BorderRight";
	static final boolean DEFAULT_BORDER_RIGHT = false;
	final SettingsModelBoolean m_borderRight =
			new SettingsModelBoolean(CFGKEY_BORDER_RIGHT, DEFAULT_BORDER_RIGHT);

	static final String CFGKEY_BORDER_INNER_VERTICAL = "BorderInnerVertical";
	static final boolean DEFAULT_BORDER_INNER_VERTICAL = false;
	final SettingsModelBoolean m_borderInnerVertical =
			new SettingsModelBoolean(CFGKEY_BORDER_INNER_VERTICAL, DEFAULT_BORDER_INNER_VERTICAL);

	static final String CFGKEY_BORDER_INNER_HORIZONTAL = "BorderInnerHorizontal";
	static final boolean DEFAULT_BORDER_INNER_HORIZONTAL = false;
	final SettingsModelBoolean m_borderInnerHorizontal =
			new SettingsModelBoolean(CFGKEY_BORDER_INNER_HORIZONTAL, DEFAULT_BORDER_INNER_HORIZONTAL);   

	/**
	 * Constructor for the node model.
	 */
	protected XlsFormatterBorderFormatterNodeModel() {

		super(
				new PortType[] { BufferedDataTable.TYPE, XlsFormatterState.TYPE_OPTIONAL },
				new PortType[] { XlsFormatterState.TYPE });
	}

	/**
	 * {@inheritDoc}
	 */
	protected PortObject[] execute(PortObject[] inObjects,
			final ExecutionContext exec) throws Exception {

		if (!XlsFormatterControlTableValidator.isControlTable((BufferedDataTable)inObjects[0],
				m_allTags.getBooleanValue() ? ControlTableType.STRING_WITHOUT_CONTENT_CHECK : ControlTableType.STRING, exec, logger))
			throw new IllegalArgumentException("The provided input table is not a valid XLS control table. See log for details.");

		XlsFormatterState xlsf = XlsFormatterState.getDeepClone(inObjects[1]);

		List<List<CellAddress>> passes;
		if (m_allTags.getBooleanValue()) {
			passes = XlsFormatterControlTableAnalysisTools.getCellsListForEachTagCombination((BufferedDataTable)inObjects[0], exec, logger);
			if (passes.size() == 0)
				setWarningMessage("Control table is empty, no tag found.");
		}
		else {
			List<CellAddress> targetCells =
					XlsFormatterControlTableAnalysisTools.getCellsMatchingTag((BufferedDataTable)inObjects[0], m_tag.getStringValue().trim(), exec, logger);
			warnOnNoMatchingTags(targetCells, m_tag.getStringValue().trim());
			passes = new ArrayList<List<CellAddress>>();
			passes.add(targetCells); // just 1 pass in this case
		}

		for (List<CellAddress> pass : passes) {
			XlsFormatterBorderFormatterLogic borderLogic = new XlsFormatterBorderFormatterLogic(
					xlsf.getCurrentSheetStateForModification().cells,
					pass,
					new XlsFormatterState.BorderEdge(
							XlsFormatterState.BorderStyle.valueOf(m_borderStyle.getStringValue().toUpperCase()),
							m_borderChangeColor.getBooleanValue() ? m_borderColor.getColorValue() : null));
	
			borderLogic.implementBordersMatchingTagsInMap(
					m_borderTop.getBooleanValue(),
					m_borderLeft.getBooleanValue(),
					m_borderBottom.getBooleanValue(),
					m_borderRight.getBooleanValue(),
					m_borderInnerVertical.getBooleanValue(),
					m_borderInnerHorizontal.getBooleanValue(),
					exec);
		}

		WarningMessageContainer warningMessageContainer = new WarningMessageContainer(); 
		XlsFormattingStateValidator.validateState(xlsf, ValidationModes.STYLES, warningMessageContainer, exec, logger);
		if (warningMessageContainer.hasMessage())
			setWarningMessage(warningMessageContainer.getMessage());
			
		return new PortObject[] { xlsf };
	}


	@Override
	protected PortObjectSpec[] configure(final PortObjectSpec[] inSpecs)
			throws InvalidSettingsException {   	

		if (!XlsFormatterControlTableValidator.isControlTableSpec((DataTableSpec)inSpecs[0], ControlTableType.STRING_WITHOUT_CONTENT_CHECK, logger))
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

		m_borderColor.setEnabled(m_borderChangeColor.getBooleanValue());

		m_tag.saveSettingsTo(settings);
		m_allTags.saveSettingsTo(settings);
		m_borderStyle.saveSettingsTo(settings);
		m_borderTop.saveSettingsTo(settings);
		m_borderBottom.saveSettingsTo(settings);
		m_borderLeft.saveSettingsTo(settings);
		m_borderRight.saveSettingsTo(settings);
		m_borderInnerVertical.saveSettingsTo(settings);
		m_borderInnerHorizontal.saveSettingsTo(settings);
		m_borderChangeColor.saveSettingsTo(settings);
		m_borderColor.saveSettingsTo(settings);

	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_tag.loadSettingsFrom(settings);
		m_allTags.loadSettingsFrom(settings);
		m_borderStyle.loadSettingsFrom(settings);
		m_borderTop.loadSettingsFrom(settings);
		m_borderBottom.loadSettingsFrom(settings);
		m_borderLeft.loadSettingsFrom(settings);
		m_borderRight.loadSettingsFrom(settings);
		m_borderInnerVertical.loadSettingsFrom(settings);
		m_borderInnerHorizontal.loadSettingsFrom(settings);
		m_borderChangeColor.loadSettingsFrom(settings);
		m_borderColor.loadSettingsFrom(settings);

	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_tag.validateSettings(settings);
		m_allTags.validateSettings(settings);
		m_borderStyle.validateSettings(settings);
		m_borderTop.validateSettings(settings);
		m_borderBottom.validateSettings(settings);
		m_borderLeft.validateSettings(settings);
		m_borderRight.validateSettings(settings);
		m_borderInnerVertical.validateSettings(settings);
		m_borderInnerHorizontal.validateSettings(settings);
		m_borderChangeColor.validateSettings(settings);
		m_borderColor.validateSettings(settings);
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
