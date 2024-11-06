/*
 * Continental Nodes for KNIME
 * Copyright (C) 2019-2021  Continental AG, Hanover, Germany
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

package com.continental.knime.xlsformatter.apply2;

import java.io.File;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.AccessDeniedException;
import java.util.EnumSet;
import java.util.Optional;

import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeLogger;
import org.knime.core.node.NodeModel;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.port.PortObject;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.util.DesktopUtil;
import org.knime.core.util.FileUtil;
import org.knime.filehandling.core.connections.FSCategory;
import org.knime.filehandling.core.connections.FSConnection;
import org.knime.filehandling.core.connections.FSFiles;
import org.knime.filehandling.core.connections.FSPath;
import org.knime.filehandling.core.connections.uriexport.URIExporter;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.ReadPathAccessor;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.SettingsModelWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.WritePathAccessor;
import org.knime.filehandling.core.defaultnodesettings.status.NodeModelStatusConsumer;
import org.knime.filehandling.core.defaultnodesettings.status.StatusMessage.MessageType;
import org.knime.filehandling.core.util.CheckNodeContextUtil;

import com.continental.knime.xlsformatter.apply.XlsFormatterApplyLogic;
import com.continental.knime.xlsformatter.commons.WarningMessageContainer;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;

final class XlsFormatterApplyNodeModel extends NodeModel {

	private static final NodeLogger LOGGER = NodeLogger.getLogger(XlsFormatterApplyNodeModel.class);
	private final WarningMessageContainer warningMessageContainer = new WarningMessageContainer();
	
	private final XlsFormatterApplySettings m_settings;

	private final int m_formatterIdx;

	private final NodeModelStatusConsumer m_statusConsumer;

	XlsFormatterApplyNodeModel(PortsConfiguration portsCfg, XlsFormatterApplySettings settings, int formatterIdx) {
		super(portsCfg.getInputPorts(), portsCfg.getOutputPorts());
		m_settings = settings;
		m_formatterIdx = formatterIdx;
		m_statusConsumer = new NodeModelStatusConsumer(EnumSet.of(MessageType.WARNING, MessageType.ERROR));
	}

	@Override
	protected PortObjectSpec[] configure(PortObjectSpec[] inSpecs) throws InvalidSettingsException {
		m_settings.getSrcFileChooser().configureInModel(inSpecs, m_statusConsumer);
		m_settings.getTgtFileChooser().configureInModel(inSpecs, m_statusConsumer);
		m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);
		return new PortObjectSpec[] {};
	}

	@Override
	protected PortObject[] execute(PortObject[] inObjects, ExecutionContext exec) throws Exception {
		final XlsFormatterState state = (XlsFormatterState) inObjects[m_formatterIdx];
		if (state.isEmpty()) {
			setWarningMessage("The XLS Formatter Port input is empty, hence nothing could be applied.");
		} else {
			try (final ReadPathAccessor readAccessor = m_settings.getSrcFileChooser().createReadPathAccessor();
					final WritePathAccessor writeAccessor = m_settings.getTgtFileChooser().createWritePathAccessor()) {
				final FSPath inputPath = readAccessor.getFSPaths(m_statusConsumer).get(0);
				final FSPath outputPath = writeAccessor.getOutputPath(m_statusConsumer);
				m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);
				createParentDirectories(outputPath);
				// since the remainder is rather costly we do this check here
				checkOverwrite(outputPath);
				m_statusConsumer.setWarningsIfRequired(this::setWarningMessage);

				// used to sneak out a warning from apply()
				XlsFormatterApplyLogic.apply(inputPath.toString(),
						() -> FSFiles.newInputStream(inputPath),
						() -> FSFiles.newOutputStream(outputPath,
								m_settings.getTgtFileChooser().getFileOverwritePolicy().getOpenOptions()),
						state,
						m_settings.getPreserveSourceNumberFormatsSettingsModel().getBooleanValue(),
						warningMessageContainer,
						exec,
						LOGGER);
				if (m_settings.getOpenOutputFileSettingsModel().getBooleanValue() && !isHeadlessOrRemote()
						&& categoryIsSupported(outputPath.toFSLocation().getFSCategory())) {
					openFile(m_settings.getTgtFileChooser(), outputPath);
				}
				if (warningMessageContainer.hasMessage()) {
					setWarningMessage(warningMessageContainer.getMessage());
				}
			}
		}
		return new PortObject[] {};
	}

	private void createParentDirectories(final FSPath outpath) throws IOException {
		final FSPath parentPath = (FSPath) outpath.getParent();
		if (parentPath != null && !FSFiles.exists(parentPath)) {
			if (m_settings.getTgtFileChooser().isCreateMissingFolders()) {
				FSFiles.createDirectories(parentPath);
			} else {
				throw new IOException(String.format(
						"The directory '%s' does not exist and must not be created due to user settings.", parentPath));
			}
		}
	}

	private void checkOverwrite(final FSPath outpath) throws AccessDeniedException, IOException {
		final FileOverwritePolicy fileOverwritePolicy = m_settings.getTgtFileChooser().getFileOverwritePolicy();
		if (fileOverwritePolicy == FileOverwritePolicy.FAIL && FSFiles.exists(outpath)) {
			throw new IOException("Output file '" + outpath.toString()
					+ "' exists and must not be overwritten due to user settings.");
		}
	}

	static boolean isHeadlessOrRemote() {
		return Boolean.getBoolean("java.awt.headless") || CheckNodeContextUtil.isRemoteWorkflowContext();
	}

	static boolean categoryIsSupported(final FSCategory fsCategory) {
		return fsCategory == FSCategory.LOCAL
				|| fsCategory == FSCategory.RELATIVE
				|| fsCategory == FSCategory.CUSTOM_URL;
	}

	private void openFile(final SettingsModelWriterFileChooser fileChooser, final FSPath outputPath)
			throws IOException {
		try (final FSConnection connection = fileChooser.getConnection()) {
			final Optional<File> file = toFile(outputPath, connection);
			if (file.isPresent()) {
				DesktopUtil.open(file.get());
			} else {
				warningMessageContainer.addMessage("Non local files cannot be opened after node execution.");
			}
		}
	}

	private static Optional<File> toFile(final FSPath outputPath, final FSConnection fsConnection) {
		final FSCategory fsCategory = outputPath.toFSLocation().getFSCategory();
		if (fsCategory == FSCategory.LOCAL) {
			return Optional.of(outputPath.toAbsolutePath().toFile());
		}
		try {
			final URIExporter uriExporter = fsConnection.getDefaultURIExporterFactory().getExporter();
			final String uri = uriExporter.toUri(outputPath).toString();
			final URL url = FileUtil.toURL(uri);
			return Optional.ofNullable(FileUtil.getFileFromURL(url));
		} catch (final MalformedURLException | IllegalArgumentException | URISyntaxException e) {
			LOGGER.debug("Unable to resolve custom URL", e);
			return Optional.empty();
		}
	}

	@Override
	protected void saveSettingsTo(NodeSettingsWO settings) {
		m_settings.saveSettingsInModel(settings);
	}

	@Override
	protected void validateSettings(NodeSettingsRO settings) throws InvalidSettingsException {
		m_settings.validateSettingsInModel(settings);
	}

	@Override
	protected void loadValidatedSettingsFrom(NodeSettingsRO settings) throws InvalidSettingsException {
		m_settings.loadSettingsInModel(settings);
	}

	@Override
	protected void loadInternals(File nodeInternDir, ExecutionMonitor exec) {
	}

	@Override
	protected void saveInternals(File nodeInternDir, ExecutionMonitor exec) {
	}

	@Override
	protected void reset() {
	}
}
