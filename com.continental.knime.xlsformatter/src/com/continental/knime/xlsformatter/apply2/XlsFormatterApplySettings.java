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

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.core.node.defaultnodesettings.SettingsModelBoolean;
import org.knime.filehandling.core.defaultnodesettings.EnumConfig;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.FileOverwritePolicy;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.SettingsModelWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filtermode.SettingsModelFilterMode.FilterMode;

final class XlsFormatterApplySettings {

	private static final String FILE_EXTENSION = ".xlsx";

	private final SettingsModelReaderFileChooser m_source;

	private final SettingsModelWriterFileChooser m_target;

	private final SettingsModelBoolean m_openOutputFile;

	XlsFormatterApplySettings(final PortsConfiguration portsCfg, final String srcGrpName, final String tgtGrpName) {
		m_source = new SettingsModelReaderFileChooser("InputFile", portsCfg, srcGrpName,
				EnumConfig.create(FilterMode.FILE), FILE_EXTENSION);
		m_target = new SettingsModelWriterFileChooser("OutputFile", portsCfg, tgtGrpName,
				EnumConfig.create(FilterMode.FILE),
				EnumConfig.create(FileOverwritePolicy.FAIL, FileOverwritePolicy.OVERWRITE), FILE_EXTENSION);
		m_openOutputFile = new SettingsModelBoolean("OpenOutputFile", false);
	}

	SettingsModelReaderFileChooser getSrcFileChooser() {
		return m_source;
	}

	SettingsModelWriterFileChooser getTgtFileChooser() {
		return m_target;
	}

	SettingsModelBoolean getOpenOutputFileSettingsModel() {
		return m_openOutputFile;
	}

	void saveSettingsInModel(final NodeSettingsWO settings) {
		m_source.saveSettingsTo(settings);
		m_target.saveSettingsTo(settings);
		m_openOutputFile.saveSettingsTo(settings);
	}

	void validateSettingsInModel(final NodeSettingsRO settings) throws InvalidSettingsException {
		m_source.validateSettings(settings);
		m_target.validateSettings(settings);
		m_openOutputFile.validateSettings(settings);
	}

	void loadSettingsInModel(final NodeSettingsRO settings) throws InvalidSettingsException {
		m_source.loadSettingsFrom(settings);
		m_target.loadSettingsFrom(settings);
		m_openOutputFile.loadSettingsFrom(settings);
	}
}
