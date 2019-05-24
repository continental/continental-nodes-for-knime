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

package com.continental.knime.xlsformatter.apply;

import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;

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
import org.knime.core.util.DesktopUtil;

import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;

public class XlsFormatterApplyNodeModel extends NodeModel {

	// the logger instance
	private static final NodeLogger logger = NodeLogger
			.getLogger(XlsFormatterApplyNodeModel.class);

	static final String CFGKEY_INPUTFILE = "InputFile";
	static final String DEFAULT_INPUTFILE = null;
	private final SettingsModelString m_inputFile =
			new SettingsModelString(XlsFormatterApplyNodeModel.CFGKEY_INPUTFILE,
					XlsFormatterApplyNodeModel.DEFAULT_INPUTFILE);

	static final String CFGKEY_OUTPUTFILE = "OutputFile";
	static final String DEFAULT_OUTPUTFILE = null;
	private final SettingsModelString m_outputFile =
			new SettingsModelString(XlsFormatterApplyNodeModel.CFGKEY_OUTPUTFILE,
					XlsFormatterApplyNodeModel.DEFAULT_OUTPUTFILE);

	static final String CFGKEY_OVERWRITE = "Overwrite";
	static final boolean DEFAULT_OVERWRITE = true;
	private final SettingsModelBoolean m_overwrite =
			new SettingsModelBoolean(XlsFormatterApplyNodeModel.CFGKEY_OVERWRITE,XlsFormatterApplyNodeModel.DEFAULT_OVERWRITE);

	static final String CFGKEY_OPENOUTPUTFILE = "OpenOutputFile";
	static final boolean DEFAULT_OPENOUTPUTFILE = false;
	private final SettingsModelBoolean m_openoutputfile =
			new SettingsModelBoolean(XlsFormatterApplyNodeModel.CFGKEY_OPENOUTPUTFILE,XlsFormatterApplyNodeModel.DEFAULT_OPENOUTPUTFILE);

	/**
	 * Constructor for the node model.
	 */
	protected XlsFormatterApplyNodeModel() {

		super(
				new PortType[] { XlsFormatterState.TYPE },
				new PortType[] {});
	}

	/**
	 * {@inheritDoc}
	 */
	protected PortObject[] execute(PortObject[] inObjects,
			final ExecutionContext exec) throws Exception {

		XlsFormatterState state = (XlsFormatterState)inObjects[0];

		String out = resolveKnimePath(m_outputFile.getStringValue().trim());
		String in = resolveKnimePath(m_inputFile.getStringValue().trim());

		if (state.isEmpty())
			setWarningMessage("The XLS Formatter Port input is empty, hence nothing could be applied.");
		else {
			
			if (!m_overwrite.getBooleanValue() && (new File(out)).isFile())
				throw new Exception("The output file already exists and overwriting is disabled in the node's configuration.");
			
			WarningMessageContainer warningMessageContainer = new WarningMessageContainer(); // used to sneak out a warning from apply()
			XlsFormatterApplyLogic.apply(in, out, null,
					state, warningMessageContainer, exec, logger);
			if (warningMessageContainer.hasMessage())
				setWarningMessage(warningMessageContainer.getMessage());
		}

		// open file:
		if (m_openoutputfile.getBooleanValue()) {
			File file = new File(out);
			if (file != null)
				DesktopUtil.open(file);
		}
		return new PortObject[] {};
	}
	
	/**
	 * Resolves a file path to a local path, esp. in regards to knime://knime.workflow/ syntax.
	 */
	private static String resolveKnimePath(String path) throws IOException, URISyntaxException {
		if (path.startsWith("knime:"))
			return org.knime.core.util.pathresolve.ResolverUtil.resolveURItoLocalFile(new URI("knime", path.substring(6), null)).getAbsolutePath();
		else
			return path;
	}

	@Override
	protected PortObjectSpec[] configure(final PortObjectSpec[] inSpecs)
			throws InvalidSettingsException {   	

		return new PortObjectSpec[] {};
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

		m_inputFile.saveSettingsTo(settings);
		m_outputFile.saveSettingsTo(settings);
		m_overwrite.saveSettingsTo(settings);
		m_openoutputfile.saveSettingsTo(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void loadValidatedSettingsFrom(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_inputFile.loadSettingsFrom(settings);
		m_outputFile.loadSettingsFrom(settings);
		m_overwrite.loadSettingsFrom(settings);
		m_openoutputfile.loadSettingsFrom(settings);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void validateSettings(final NodeSettingsRO settings)
			throws InvalidSettingsException {

		m_inputFile.validateSettings(settings);
		m_outputFile.validateSettings(settings);
		m_overwrite.validateSettings(settings);
		m_openoutputfile.validateSettings(settings);
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
