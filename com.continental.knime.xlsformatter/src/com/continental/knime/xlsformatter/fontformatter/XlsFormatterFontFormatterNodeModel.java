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

package com.continental.knime.xlsformatter.fontformatter;

import java.awt.Color;
import java.io.File;
import java.io.IOException;
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
import org.knime.core.node.defaultnodesettings.SettingsModelInteger;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.apply.XlsFormatterApplyLogic;
import com.continental.knime.xlsformatter.commons.AddressingTools;
import com.continental.knime.xlsformatter.commons.TagBasedXlsCellFormatterNodeModel;
import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableAnalysisTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;
import com.continental.knime.xlsformatter.commons.XlsFormatterTagTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.FormattingFlag;
import com.continental.knime.xlsformatter.porttype.XlsFormatterStateSpec;

public class XlsFormatterFontFormatterNodeModel extends TagBasedXlsCellFormatterNodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsFormatterFontFormatterNodeModel.class);

	//Input tag string
	static final String CFGKEY_TAGSTRING = "Tag";
	static final String DEFAULT_TAGSTRING = "header";
	final SettingsModelString m_tagstring =
			new SettingsModelString(CFGKEY_TAGSTRING, DEFAULT_TAGSTRING);    

	static final String CFGKEY_BOLD = "Bold";
	static final boolean DEFAULT_BOLD = false;
	final SettingsModelBoolean m_bold =
			new SettingsModelBoolean(CFGKEY_BOLD, DEFAULT_BOLD);  

	static final String CFGKEY_ITALIC = "Italic";
	static final boolean DEFAULT_ITALIC = false;
	final SettingsModelBoolean m_italic =
			new SettingsModelBoolean(CFGKEY_ITALIC, DEFAULT_ITALIC);

	static final String CFGKEY_UNDERLINE = "Underline";
	static final boolean DEFAULT_UNDERLINE = false;
	final SettingsModelBoolean m_underline =
			new SettingsModelBoolean(CFGKEY_UNDERLINE, DEFAULT_UNDERLINE);  

	static final String CFGKEY_CHANGESIZE = "ChangeSize";
	static final boolean DEFAULT_CHANGESIZE = false;
	final SettingsModelBoolean m_changesize =
			new SettingsModelBoolean(CFGKEY_CHANGESIZE, DEFAULT_CHANGESIZE);  

	static final String CFGKEY_SIZE = "Size";
	static final Integer DEFAULT_SIZE = 12;
	final SettingsModelInteger m_size =
			new SettingsModelInteger(CFGKEY_SIZE, DEFAULT_SIZE);  

	static final String CFGKEY_CHANGECOLOR = "ChangeColor";
	static final boolean DEFAULT_CHANGECOLOR = false;
	final SettingsModelBoolean m_changecolor =
			new SettingsModelBoolean(CFGKEY_CHANGECOLOR, DEFAULT_CHANGECOLOR);  

	static final String CFGKEY_FONTCOLOR = "FontColor";
	static final Color DEFAULT_FONTCOLOR = new Color(0,0,0);
	final SettingsModelColor m_fontcolor =
			new SettingsModelColor(CFGKEY_FONTCOLOR, DEFAULT_FONTCOLOR);


	/**
	 * Constructor for the node model.
	 */
	protected XlsFormatterFontFormatterNodeModel() {
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

		for (CellAddress cell : targetCells) {
			XlsFormatterState.CellState cellState = AddressingTools.safelyGetCellInMap(xlsf.cells, cell);

			FormattingFlag flag = XlsFormatterUiOptions.getFormattingFlagFromBoolean(m_bold.getBooleanValue());
			if (flag != FormattingFlag.UNMODIFIED)
				cellState.fontBold = flag;

			flag = XlsFormatterUiOptions.getFormattingFlagFromBoolean(m_italic.getBooleanValue());
			if (flag != FormattingFlag.UNMODIFIED)
				cellState.fontItalic = flag;

			flag = XlsFormatterUiOptions.getFormattingFlagFromBoolean(m_underline.getBooleanValue());
			if (flag != FormattingFlag.UNMODIFIED)
				cellState.fontUnderline = flag;

			if (m_changesize.getBooleanValue())
				cellState.fontSize = m_size.getIntValue();

			if (m_changecolor.getBooleanValue())
				cellState.fontColor = m_fontcolor.getColorValue();
		}

		WarningMessageContainer warningMessageContainer = new WarningMessageContainer(); 
		XlsFormatterApplyLogic.checkDerivedStyleComplexity(xlsf, warningMessageContainer, exec, logger);
		if (warningMessageContainer.hasMessage())
			setWarningMessage(warningMessageContainer.getMessage());

		return new PortObject[] { xlsf };
	}


	@Override
	protected PortObjectSpec[] configure(final PortObjectSpec[] inSpecs)
			throws InvalidSettingsException {   	

		if (m_tagstring.getStringValue().trim().equals(""))
			throw new IllegalArgumentException("Empty tag is not allowed, you need to enter something here that is also present in your control table (e.g., \"x\" or \"data\"");

		if (!XlsFormatterTagTools.isValidSingleTag(m_tagstring.getStringValue().trim()))
			throw new IllegalArgumentException("Only a single tag is allowed, no comma-separated list.");

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

		m_size.setEnabled(m_changesize.getBooleanValue());
		m_fontcolor.setEnabled(m_changecolor.getBooleanValue());

		m_tagstring.saveSettingsTo(settings);
		m_bold.saveSettingsTo(settings);
		m_italic.saveSettingsTo(settings);
		m_underline.saveSettingsTo(settings);
		m_size.saveSettingsTo(settings);
		m_fontcolor.saveSettingsTo(settings);
		m_changesize.saveSettingsTo(settings);
		m_changecolor.saveSettingsTo(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_tagstring.loadSettingsFrom(settings);
		m_bold.loadSettingsFrom(settings);
		m_italic.loadSettingsFrom(settings);
		m_underline.loadSettingsFrom(settings);
		m_size.loadSettingsFrom(settings);
		m_fontcolor.loadSettingsFrom(settings);
		m_changesize.loadSettingsFrom(settings);
		m_changecolor.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_tagstring.validateSettings(settings);
		m_bold.validateSettings(settings);
		m_italic.validateSettings(settings);
		m_underline.validateSettings(settings);
		m_size.validateSettings(settings);
		m_fontcolor.validateSettings(settings);
		m_changesize.validateSettings(settings);
		m_changecolor.validateSettings(settings);
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
