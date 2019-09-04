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

package com.continental.knime.xlsformatter.xlscontroltablegenerator;

import java.io.File;
import java.io.IOException;

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
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.commons.XlsFormatterControlTableAnalysisTools;
import com.continental.knime.xlsformatter.commons.XlsFormatterUiOptions;

public class XlsControlTableGeneratorNodeModel extends NodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsControlTableGeneratorNodeModel.class);

	static final String CFGKEY_ROW_SHIFT = "RowShift";
	static final boolean DEFAULT_ROW_SHIFT = false;
	final SettingsModelBoolean m_rowShift =
			new SettingsModelBoolean(CFGKEY_ROW_SHIFT, DEFAULT_ROW_SHIFT);

	static final String CFGKEY_UNPIVOT = "Unpivot";
	static final boolean DEFAULT_UNPIVOT = false;
	final SettingsModelBoolean m_unpivot =
			new SettingsModelBoolean(CFGKEY_UNPIVOT, DEFAULT_UNPIVOT);

	static final String CFGKEY_EXTENDED_UNPIVOT_COLUMNS = "ExtendedHeaderColumns";
	static final boolean DEFAULT_EXTENDED_UNPIVOT_COLUMNS = true;
	final SettingsModelBoolean m_extendedColumnsAtUnpivot =
			new SettingsModelBoolean(CFGKEY_EXTENDED_UNPIVOT_COLUMNS, DEFAULT_EXTENDED_UNPIVOT_COLUMNS);

	static final String CFGKEY_INCONSISTENCY_RESOLUTION_STRATEGY = "InconsistencyResolutionStrategy";
	static final String DEFAULT_INCONSISTENCY_RESOLUTION_STRATEGY = InconsistencyResolutionOptions.FAIL.toString();
	final SettingsModelString m_inconsistencyResolutionStrategy =
			new SettingsModelString(CFGKEY_INCONSISTENCY_RESOLUTION_STRATEGY, DEFAULT_INCONSISTENCY_RESOLUTION_STRATEGY);
	
	static final String CFGKEY_OPERATION_TYPE = "OperationType";
	static final String DEFAULT_OPERATION_TYPE = OperationType.STANDARD.toString(); // no real default, as it will be automatically overwritten based on incoming table
	final SettingsModelString m_operationType =
			new SettingsModelString(CFGKEY_OPERATION_TYPE, DEFAULT_OPERATION_TYPE);
	
	
	/**
	 * The type of operation that this node shall perform (defined via the layout of the
	 * incoming data table).
	 */
	public enum OperationType {
		STANDARD,
		PIVOT_BACK;
		
		@Override
		public String toString() {
			switch (this) {
			case STANDARD:
				return "from arbitrary input table to XLS Control Table (wide or long/unpivoted)";
			case PIVOT_BACK:
				return "from long/unpivoted layout to wide XLS Control Table";
			default:
				return this.toString().toLowerCase();
			}
		}
		
		public static OperationType getFromString(String value) {
	    return XlsFormatterUiOptions.getEnumEntryFromString(OperationType.values(), value);
		}
	}
	
	/**
	 * In case of an input table that was created with the 'unpivoted with extra columns' option, redundant information
	 * is contained in this input table in regards to columns. For example, the numeric column index might have been shifted
	 * with a Math Formula node. In the manual alternative of Pivoting and RowID, the user could have defined that this
	 * specific 'Column (number)' node would have been used as the new column headers (i.e. via an additional workaround by
	 * joining in the corresponding column letters, to be precise). In fact, this node now sees multiple columns holding
	 * inconsistent values regarding the target column of a data row. Option FAIL leads to an exception being thrown in case
	 * an inconsistency is detected. The other four options determine which of the input columns wins (with the partially
	 * inconsistent information in the other three columns being ignored).
	 */
	public enum InconsistencyResolutionOptions {
		FAIL,
		CELL,
		COLUMN,
		COLUMN_COMPARABLE,
		COLUMN_NUMBER;
		
		@Override
		public String toString() {
			switch (this) {
			case FAIL:
				return "fail";
			case CELL:
				return "use 'Cell'";
			case COLUMN:
				return "use 'Column' and 'Row'";
			case COLUMN_COMPARABLE:
				return "use 'Column (comparable)' and 'Row'";
			case COLUMN_NUMBER:
				return "use 'Column (number)' and 'Row'";
			default:
				return this.toString().toLowerCase();
			}
		}
		
		public static InconsistencyResolutionOptions getFromString(String value) {
	    return XlsFormatterUiOptions.getEnumEntryFromString(InconsistencyResolutionOptions.values(), value);
		}
	}
	
	/**
	 * Constructor for the node model.
	 */
	protected XlsControlTableGeneratorNodeModel() {

		super(
				new PortType[] { BufferedDataTable.TYPE },
				new PortType[] { BufferedDataTable.TYPE });
	}

	/**
	 * {@inheritDoc}
	 */
	protected BufferedDataTable[] execute(BufferedDataTable[] inData,
			final ExecutionContext exec) throws Exception {

		WarningMessageContainer warningMessageContainer = new WarningMessageContainer();
		BufferedDataTable[] ret = null;
		
		XlsFormatterControlTableAnalysisTools.isLongControlTableSpec(inData[0].getSpec(), warningMessageContainer, logger); // just to add the warning about potentially intended pivot-back option
		
		if (m_operationType.getStringValue().equals(OperationType.PIVOT_BACK.toString()))
			ret = XlsControlTableGeneratorFunctionPivotBack.pivotBack(
					inData[0], warningMessageContainer,
					InconsistencyResolutionOptions.getFromString(m_inconsistencyResolutionStrategy.getStringValue()) , exec, logger);
		
		else if (m_unpivot.getBooleanValue())
			ret = XlsControlTableGeneratorFunctionUnpivot.unpivot(inData[0], m_rowShift.getBooleanValue(), exec, logger, m_extendedColumnsAtUnpivot.getBooleanValue());

		else
			ret = XlsControlTableGeneratorFunctionDerivePivoted.derivePivoted(
					inData[0], m_rowShift.getBooleanValue(), true, warningMessageContainer, exec, logger);
		
		if (warningMessageContainer.hasMessage())
			setWarningMessage(warningMessageContainer.getMessage());
		return ret;
	}

	@Override
	protected DataTableSpec[] configure(final DataTableSpec[] inSpecs)
			throws InvalidSettingsException {

		WarningMessageContainer warningMessageContainer = new WarningMessageContainer();
		
		int longUnpivotedInputTableColumnCount = XlsFormatterControlTableAnalysisTools.isLongControlTableSpec(inSpecs[0], warningMessageContainer, logger);
		m_operationType.setStringValue(longUnpivotedInputTableColumnCount == -1 ?
				OperationType.STANDARD.toString() : OperationType.PIVOT_BACK.toString());
		
		if (longUnpivotedInputTableColumnCount != 8)
			m_inconsistencyResolutionStrategy.setStringValue(InconsistencyResolutionOptions.FAIL.toString());
		
		if (warningMessageContainer.hasMessage())
			setWarningMessage(warningMessageContainer.getMessage());
		
		if (m_operationType.getStringValue().equals(OperationType.PIVOT_BACK.toString()))
			return new DataTableSpec[] { null }; // when pivoting back, we cannot know the spec as it depends on the data table's content, which is not known yet at the configuration step
		
		if (m_unpivot.getBooleanValue())
			return new DataTableSpec[] { XlsControlTableGeneratorFunctionUnpivot.getUnpivotSpec(inSpecs[0], m_extendedColumnsAtUnpivot.getBooleanValue()) };

		return new DataTableSpec[] { XlsControlTableGeneratorFunctionDerivePivoted.getPivotedSpec(inSpecs[0]) };
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

		m_extendedColumnsAtUnpivot.setEnabled(m_unpivot.getBooleanValue());
		
		m_rowShift.saveSettingsTo(settings);
		m_unpivot.saveSettingsTo(settings);
		m_extendedColumnsAtUnpivot.saveSettingsTo(settings);
		m_inconsistencyResolutionStrategy.saveSettingsTo(settings);
		m_operationType.saveSettingsTo(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_rowShift.loadSettingsFrom(settings);
		m_unpivot.loadSettingsFrom(settings);
		m_extendedColumnsAtUnpivot.loadSettingsFrom(settings);
		if (settings.containsKey(CFGKEY_OPERATION_TYPE))
			m_operationType.loadSettingsFrom(settings);
		if (settings.containsKey(CFGKEY_INCONSISTENCY_RESOLUTION_STRATEGY))
			m_inconsistencyResolutionStrategy.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_rowShift.validateSettings(settings);
		m_unpivot.validateSettings(settings);
		m_extendedColumnsAtUnpivot.validateSettings(settings);
		if (settings.containsKey(CFGKEY_OPERATION_TYPE))
			m_operationType.validateSettings(settings);
		if (settings.containsKey(CFGKEY_INCONSISTENCY_RESOLUTION_STRATEGY))
			m_inconsistencyResolutionStrategy.validateSettings(settings);
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
