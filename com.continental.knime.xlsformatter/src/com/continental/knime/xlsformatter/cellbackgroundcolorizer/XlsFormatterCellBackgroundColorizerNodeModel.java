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

package com.continental.knime.xlsformatter.cellbackgroundcolorizer;

import java.awt.Color;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

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

import com.continental.knime.xlsformatter.apply.XlsFormatterApplyLogic;
import com.continental.knime.xlsformatter.commons.AddressingTools;
import com.continental.knime.xlsformatter.commons.ColorTools;
import com.continental.knime.xlsformatter.commons.TagBasedXlsCellFormatterNodeModel;
import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableAnalysisTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator.ControlTableType;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.FillPattern;
import com.continental.knime.xlsformatter.porttype.XlsFormatterStateSpec;


public class XlsFormatterCellBackgroundColorizerNodeModel extends TagBasedXlsCellFormatterNodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsFormatterCellBackgroundColorizerNodeModel.class);

	//Control Table Selection Radion Button
	static final String CFGKEY_CONTROLTABLESTYLE = XlsFormatterUiOptions.UI_LABEL_CONTROLTABLESTYLE_KEY;
	static final String OPTION_CONTROLTABLESTYLE_STANDARD = XlsFormatterUiOptions.UI_LABEL_CONTROLTABLESTYLE_STANDARD;
	static final String OPTION_CONTROLTABLESTYLE_DIRECT = "direct color codes in RGB format";
	static final String DEFAULT_CONTROLTABLESTYLE = OPTION_CONTROLTABLESTYLE_STANDARD;
	final SettingsModelString m_controltablestyle =
			new SettingsModelString(CFGKEY_CONTROLTABLESTYLE, DEFAULT_CONTROLTABLESTYLE);   

	//Input tag string
	static final String CFGKEY_TAGSTRING = "Tag";
	static final String DEFAULT_TAGSTRING = "header";
	final SettingsModelString m_tagstring =
			new SettingsModelString(CFGKEY_TAGSTRING, DEFAULT_TAGSTRING);

	//Background color selection
	static final String CFGKEY_BACKGROUND_COLOR_SELECTION = "BackGroundColorSelection";
	static final boolean DEFAULT_BACKGROUND_COLOR_SELECTION = true;
	final SettingsModelBoolean m_backgroundcolorselection =
			new SettingsModelBoolean(CFGKEY_BACKGROUND_COLOR_SELECTION, DEFAULT_BACKGROUND_COLOR_SELECTION);

	//Background color field
	static final String CFGKEY_BACKGROUNDCOLOR = "BackGroundColor";
	static final Color DEFAULT_BACKGROUNDCOLOR = new Color(255, 255, 0);
	final SettingsModelColor m_backgroundcolor =
			new SettingsModelColor(CFGKEY_BACKGROUNDCOLOR, DEFAULT_BACKGROUNDCOLOR);

	//Background pattern selection
	static final String CFGKEY_BACKGROUND_PATTERN_SELECTION = "BackGroundPatternSelection";
	static final String DEFAULT_BACKGROUND_PATTERN_SELECTION = XlsFormatterState.FillPattern.UNMODIFIED.toString().toLowerCase();
	final SettingsModelString m_backgroundpatternselection =
			new SettingsModelString(CFGKEY_BACKGROUND_PATTERN_SELECTION, DEFAULT_BACKGROUND_PATTERN_SELECTION);

	//Background pattern fill dropdwon
	static final String CFGKEY_BACKGROUND_PATTERN_COLOR_SELECTION = "BackGroundPatternColorSelection";
	static final boolean DEFAULT_BACKGROUND_PATTERN_COLOR_SELECTION = false;
	final SettingsModelBoolean m_backgroundpatterncolorselection =
			new SettingsModelBoolean(CFGKEY_BACKGROUND_PATTERN_COLOR_SELECTION, DEFAULT_BACKGROUND_PATTERN_COLOR_SELECTION);

	//Background pattern color
	static final String CFGKEY_BACKGROUND_PATTERN_COLOR = "BackGroundPatternColor";
	static final Color DEFAULT_BACKGROUND_PATTERN_COLOR = new Color(0, 0, 0);
	final SettingsModelColor m_backgroundpatterncolor =
			new SettingsModelColor(CFGKEY_BACKGROUND_PATTERN_COLOR, DEFAULT_BACKGROUND_PATTERN_COLOR);


	public static List<String> fillPatternDropdownOptions = null;

	static {
		fillPatternDropdownOptions = XlsFormatterUiOptions.getDropdownListFromEnum(XlsFormatterState.FillPattern.values());
		fillPatternDropdownOptions.remove(XlsFormatterState.FillPattern.NONE.toString().toLowerCase()); // because we currently don't support "undo" in the UI
		fillPatternDropdownOptions.remove(XlsFormatterState.FillPattern.SOLID_BACKGROUND_COLOR.toString().toLowerCase()); // because we handle this logically different in our UI than in the XlsFormatterState and POI
	}

	/**
	 * Constructor for the node model.
	 */
	protected XlsFormatterCellBackgroundColorizerNodeModel() {

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
				m_controltablestyle.getStringValue().equals(OPTION_CONTROLTABLESTYLE_STANDARD) ? ControlTableType.STRING : ControlTableType.STRING_WITHOUT_CONTENT_CHECK, exec, logger))
			throw new IllegalArgumentException("The provided input table is not a valid XLS control table. See log for details.");
		

		XlsFormatterState xlsf = XlsFormatterState.getDeepClone(inObjects[1]);

		if (m_controltablestyle.getStringValue().equals(OPTION_CONTROLTABLESTYLE_STANDARD)) {
			List<CellAddress> targetCells =
					XlsFormatterControlTableAnalysisTools.getCellsMatchingTag((BufferedDataTable)inObjects[0], m_tagstring.getStringValue().trim(), exec, logger);
			warnOnNoMatchingTags(targetCells, m_tagstring.getStringValue().trim());

			for (CellAddress cell : targetCells) {
				XlsFormatterState.CellState cellState = AddressingTools.safelyGetCellInMap(xlsf.cells, cell);

				FillPattern fillPatternSelection = FillPattern.valueOf(m_backgroundpatternselection.getStringValue().toUpperCase());

				if (m_backgroundcolorselection.getBooleanValue()) {
					cellState.backgroundColor = m_backgroundcolor.getColorValue();
					if (
							(cellState.fillPattern == FillPattern.UNMODIFIED || cellState.fillPattern == FillPattern.NONE) &&
							fillPatternSelection == FillPattern.UNMODIFIED)
						cellState.fillPattern = FillPattern.SOLID_BACKGROUND_COLOR;
				}

				if (fillPatternSelection != FillPattern.UNMODIFIED)
					cellState.fillPattern = fillPatternSelection;
				if (m_backgroundpatterncolorselection.getBooleanValue())
					cellState.fillForegroundColor = m_backgroundpatterncolor.getColorValue();
			}
		}
		else { // option direct RGB values in control table instead of tags
			Map<CellAddress, String> cellContentMap =
					XlsFormatterControlTableAnalysisTools.getCellsWithStringValues((BufferedDataTable)inObjects[0], exec, logger);

			for (CellAddress cell : cellContentMap.keySet()) {
				XlsFormatterState.CellState cellState;
				if (xlsf.cells.containsKey(cell))
					cellState = xlsf.cells.get(cell);
				else {
					cellState = new XlsFormatterState.CellState();
					xlsf.cells.put(cell, cellState);
				}

				String value = cellContentMap.get(cell);
				if (!value.trim().equals("")) {
					cellState.backgroundColor = ColorTools.anyColorCodeToXlsfColor(value);
					cellState.fillPattern = FillPattern.SOLID_BACKGROUND_COLOR;
				}
			}
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

		m_backgroundcolor.setEnabled(true);
		m_backgroundpatterncolor.setEnabled(false);    	

		m_controltablestyle.saveSettingsTo(settings);
		m_tagstring.saveSettingsTo(settings);
		m_backgroundcolor.saveSettingsTo(settings);
		m_backgroundpatterncolorselection.saveSettingsTo(settings);
		m_backgroundpatterncolor.saveSettingsTo(settings);
		m_backgroundpatternselection.saveSettingsTo(settings);
		m_backgroundcolorselection.saveSettingsTo(settings);

	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_controltablestyle.loadSettingsFrom(settings);
		m_tagstring.loadSettingsFrom(settings);
		m_backgroundcolor.loadSettingsFrom(settings);
		m_backgroundpatterncolorselection.loadSettingsFrom(settings);
		m_backgroundpatterncolor.loadSettingsFrom(settings);
		m_backgroundpatternselection.loadSettingsFrom(settings);
		m_backgroundcolorselection.loadSettingsFrom(settings);

	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_controltablestyle.validateSettings(settings);
		m_tagstring.validateSettings(settings);
		m_backgroundcolor.validateSettings(settings);
		m_backgroundpatterncolorselection.validateSettings(settings);
		m_backgroundpatterncolor.validateSettings(settings);
		m_backgroundpatternselection.validateSettings(settings);
		m_backgroundcolorselection.validateSettings(settings);


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
