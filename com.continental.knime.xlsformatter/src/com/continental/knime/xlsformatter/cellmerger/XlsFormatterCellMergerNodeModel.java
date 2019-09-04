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

package com.continental.knime.xlsformatter.cellmerger;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
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
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.commons.AddressingTools;
import com.continental.knime.xlsformatter.commons.TagBasedXlsCellFormatterNodeModel;
import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableAnalysisTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator.ControlTableType;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterStateSpec;

public class XlsFormatterCellMergerNodeModel extends TagBasedXlsCellFormatterNodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsFormatterCellMergerNodeModel.class);

	//Input tag string
	static final String CFGKEY_TAGSTRING = "Tag(s)";
	static final String DEFAULT_TAGSTRING = "header";
	final SettingsModelString m_tag =
			new SettingsModelString(CFGKEY_TAGSTRING,DEFAULT_TAGSTRING);
	
	static final String CFGKEY_ALL_TAGS = "AllTags";
	static final boolean DEFAULT_ALL_TAGS = false;
	final SettingsModelBoolean m_allTags =
			new SettingsModelBoolean(CFGKEY_ALL_TAGS, DEFAULT_ALL_TAGS);


	/**
	 * Constructor for the node model.
	 */
	protected XlsFormatterCellMergerNodeModel() {

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
		XlsFormatterState.SheetState xlsfs = xlsf.getCurrentSheetStateForModification();
		
		// define list of to merge tags
		List<String> tagsToMerge = new ArrayList<String>();
		if (m_allTags.getBooleanValue())
			tagsToMerge = XlsFormatterControlTableAnalysisTools.getAllFullCellContentTags((BufferedDataTable)inObjects[0], exec, logger); 
		else
			tagsToMerge.add(m_tag.getStringValue());
		
		WarningMessageContainer warningMessage = new WarningMessageContainer();
		
		// derive list of ranges from list of tags and fail in case of inconsistencies:
		List<CellRangeAddress> newRanges = new ArrayList<CellRangeAddress>(); 
		for (String tag : tagsToMerge) {
			List<CellRangeAddress> mergeRanges = XlsFormatterControlTableAnalysisTools.getRangesFromTag((BufferedDataTable)inObjects[0], tag, m_allTags.getBooleanValue(), 
					false, warningMessage, exec, logger);
			if (mergeRanges.size() != 0)
				logger.debug("Detected merge ranges for tag \"" + tag + "\" as: " + mergeRanges.stream().map(r -> r.formatAsString()).collect(Collectors.joining(";")));
			if (AddressingTools.hasOverlap(newRanges, mergeRanges, exec, logger))
				throw new IllegalArgumentException("The provided merge ranges overlap with eachother. See log for details.");
			newRanges.addAll(mergeRanges);
		}
		if (AddressingTools.hasOverlap(xlsf.getCurrentSheetStateForModification().mergeRanges, newRanges, exec, logger))
			throw new IllegalArgumentException("The provided merge ranges overlap with previously defined ranges. See log for details.");
		
		// check whether an individual merge range will loose prior defined varying formatting instructions
		// of the contained cells (except border formatting) and warn: 
		boolean showWarningAboutMergeFormattingLoss = false;
		for (CellRangeAddress mergeRange : newRanges) {
			String topLeftCellFormat = null;
			CellAddress topLeftCellAddress = new CellAddress(mergeRange.getFirstRow(), mergeRange.getFirstColumn());
			for (int r = mergeRange.getFirstRow(); r <= mergeRange.getLastRow(); r++) {
				for (int c = mergeRange.getFirstColumn(); c <= mergeRange.getLastColumn(); c++) {
					CellAddress cellAddress = new CellAddress(r, c);
					CellState cellState = xlsfs.cells.containsKey(cellAddress) ? xlsfs.cells.get(cellAddress) : null;
					String cellShortString = cellState == null ? "[na]" : cellState.cellFormatToShortString(false, false);
					if (topLeftCellFormat == null) // this means we are at the beginning of the loop, i.e. at the top-left cell
						topLeftCellFormat = cellShortString;
					else // topLeftCellFormat was already set, we now search for a deviation from it
						if (cellState != null && !cellShortString.equals(topLeftCellFormat)) {
							showWarningAboutMergeFormattingLoss = true;
							logger.debug("For merge range " + mergeRange.formatAsString() + ", cell " + cellAddress.formatAsString() + " has different formatting instructions (not considering border formats) than the top-left cell " + topLeftCellAddress.formatAsString() + ".");
						}
				}
			}
		}
		if (showWarningAboutMergeFormattingLoss)
			warningMessage.addMessage("This merge makes prior cell formatting instructions obsolete as only the top-left cell of a merged range will be considered by the XLS Formatter (apply) node.");
		
		// show warnings and implement merge in state:
		if (warningMessage.hasMessage())
			setWarningMessage(warningMessage.getMessage());
		xlsfs.mergeRanges.addAll(newRanges);

		return new PortObject[] { xlsf };
	}

	@Override
	protected PortObjectSpec[] configure(final PortObjectSpec[] inSpecs)
			throws InvalidSettingsException {   	

		if (!XlsFormatterControlTableValidator.isControlTableSpec((DataTableSpec)inSpecs[0], logger))
			throw new InvalidSettingsException("The configured input table header is not that of a valid XLS Formatting control table. See log for details.");

		XlsFormatterStateSpec spec = inSpecs[1] == null ? XlsFormatterStateSpec.getEmptySpec() : ((XlsFormatterStateSpec)inSpecs[1]).getCopy();
		
		return new PortObjectSpec[] { spec };
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void saveSettingsTo(final NodeSettingsWO settings) {

		m_tag.saveSettingsTo(settings);
		m_allTags.saveSettingsTo(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_tag.loadSettingsFrom(settings);
		m_allTags.loadSettingsFrom(settings);

	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_tag.validateSettings(settings);
		m_allTags.validateSettings(settings);

	}

	@Override
	protected void loadInternals(File nodeInternDir, ExecutionMonitor exec)
			throws IOException, CanceledExecutionException {
	}

	@Override
	protected void saveInternals(File nodeInternDir, ExecutionMonitor exec)
			throws IOException, CanceledExecutionException {
	}

	@Override
	protected void reset() {
	}
}
