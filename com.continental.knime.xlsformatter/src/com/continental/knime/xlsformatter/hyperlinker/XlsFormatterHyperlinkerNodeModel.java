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

package com.continental.knime.xlsformatter.hyperlinker;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.util.CellAddress;
import org.knime.core.data.DataTableSpec;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.CanceledExecutionException;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableAnalysisTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator;
import com.continental.knime.xlsformatter.commons.XlsFormattingStateValidator;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableValidator.ControlTableType;
import com.continental.knime.xlsformatter.commons.XlsFormattingStateValidator.ValidationModes;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.CellState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterStateSpec;

public class XlsFormatterHyperlinkerNodeModel extends NodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsFormatterHyperlinkerNodeModel.class);


	/**
	 * Constructor for the node model.
	 */
	protected XlsFormatterHyperlinkerNodeModel() {
		super(
				new PortType[] { BufferedDataTable.TYPE, XlsFormatterState.TYPE_OPTIONAL },
				new PortType[] { XlsFormatterState.TYPE });
	}

	/**
	 * {@inheritDoc}
	 */
	protected PortObject[] execute(PortObject[] inObjects,
			final ExecutionContext exec) throws Exception {
		
		if (!XlsFormatterControlTableValidator.isControlTable((BufferedDataTable)inObjects[0], ControlTableType.STRING_WITHOUT_CONTENT_CHECK, exec, logger))
			throw new IllegalArgumentException("The provided input table is not a valid XLS control table. See log for details.");

		XlsFormatterState xlsf = XlsFormatterState.getDeepClone(inObjects[1]);
		XlsFormatterState.SheetState xlsfs = xlsf.getCurrentSheetStateForModification();
		WarningMessageContainer warningMessageContainer = new WarningMessageContainer();

		Map<CellAddress, String> hyperlinkMap =
				XlsFormatterControlTableAnalysisTools.getCellStringMaps((BufferedDataTable)inObjects[0], false, exec, logger).cellAddressToStringValueMap;
		
		if (hyperlinkMap.size() == 0)
			warningMessageContainer.addMessage("No hyperlinks found in input table.");
		else {
			
			// check for a partly overlap of these target cells with a previously merged range and warn
			String mergeOverlapRanges = XlsFormatterControlTableAnalysisTools.getOverlappingRanges(new ArrayList<CellAddress>(hyperlinkMap.keySet()), xlsfs.mergeRanges, exec, logger);
			if (mergeOverlapRanges != null)
				warningMessageContainer.addMessage("Modification on parts of previously merged range(s) (" + mergeOverlapRanges + ") will have no effect.");
		}

		XlsFormatterHyperlinkerPreCheck.validateHyperlinks(hyperlinkMap.entrySet());
		
		for (Entry<CellAddress, String> entry : hyperlinkMap.entrySet()) {
			CellState cellState = null;
			if (xlsfs.cells.containsKey(entry.getKey()))
				cellState = xlsfs.cells.get(entry.getKey());
			else {
				cellState = new CellState();
				xlsfs.cells.put(entry.getKey(), cellState);
			}
			
			cellState.hyperlink = entry.getValue();
		}
		
		XlsFormattingStateValidator.validateState(xlsf, ValidationModes.LINKS, warningMessageContainer, exec, logger);
		
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

	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

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
