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

import java.awt.Component;
import java.awt.GridBagLayout;

import javax.swing.BorderFactory;
import javax.swing.JLabel;
import javax.swing.JPanel;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NodeSettingsWO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.defaultnodesettings.DialogComponentBoolean;
import org.knime.core.node.port.PortObjectSpec;
import org.knime.core.node.util.SharedIcons;
import org.knime.filehandling.core.data.location.variable.FSLocationVariableType;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.DialogComponentReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.reader.SettingsModelReaderFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.DialogComponentWriterFileChooser;
import org.knime.filehandling.core.defaultnodesettings.filechooser.writer.SettingsModelWriterFileChooser;
import org.knime.filehandling.core.util.GBCBuilder;

final class XlsFormatterApplyNodeDialog extends NodeDialogPane {

	private final DialogComponentReaderFileChooser m_source;

	private final DialogComponentWriterFileChooser m_target;

	private final DialogComponentBoolean m_openOutputFile;
	private final JLabel m_openOutputFileLbl;
	
	private final DialogComponentBoolean m_preserveSourceNumberFormats;
	

	XlsFormatterApplyNodeDialog(final XlsFormatterApplySettings settings) {
		final SettingsModelReaderFileChooser srcFileChooser = settings.getSrcFileChooser();
		m_source = new DialogComponentReaderFileChooser(srcFileChooser, "xls_apply",
				createFlowVariableModel(srcFileChooser.getKeysForFSLocation(), FSLocationVariableType.INSTANCE));
		final SettingsModelWriterFileChooser tgtFileChooser = settings.getTgtFileChooser();
		m_target = new DialogComponentWriterFileChooser(tgtFileChooser, "xls_apply",
				createFlowVariableModel(tgtFileChooser.getKeysForFSLocation(), FSLocationVariableType.INSTANCE));
		m_preserveSourceNumberFormats = new DialogComponentBoolean(settings.getPreserveSourceNumberFormatsSettingsModel(),
				"preserve source file's cell number formats (e.g. date cells written by KNIME)");
		m_openOutputFile = new DialogComponentBoolean(settings.getOpenOutputFileSettingsModel(),
				"open output file after execution");
		m_openOutputFileLbl = new JLabel("");

		m_target.getModel().addChangeListener(l -> toggleOpenFileAfterExecOption());
		addTab("Settings", createSettingsPanel());

	}

	private Component createSettingsPanel() {
		final JPanel p = new JPanel(new GridBagLayout());
		GBCBuilder gbc = new GBCBuilder().anchorLineStart().weight(1, 0).resetPos().setWidth(2).fillHorizontal()
				.insetLeft(5);
		p.add(createSourcePanel(), gbc.build());
		p.add(createDestinationPanel(), gbc.incY().build());
		p.add(m_preserveSourceNumberFormats.getComponentPanel(),
				gbc.incY().setWidth(1).setWeightX(0).fillNone().insetLeft(0).build());
		p.add(m_openOutputFile.getComponentPanel(),
				gbc.incY().setWidth(1).setWeightX(0).fillNone().insetLeft(0).build());
		p.add(m_openOutputFileLbl, gbc.incX().insetLeft(5).build());
		p.add(new JPanel(), gbc.incY().setWidth(2).weight(1, 1).fillBoth().build());
		return p;
	}

	private JPanel createSourcePanel() {
		final JPanel panel = new JPanel(new GridBagLayout());
		final GBCBuilder gbc = new GBCBuilder().anchorLineStart().weight(1, 0).resetPos().fillHorizontal();
		panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Source"));
		panel.add(m_source.getComponentPanel(), gbc.build());
		return panel;
	}

	private JPanel createDestinationPanel() {
		final JPanel panel = new JPanel(new GridBagLayout());
		final GBCBuilder gbc = new GBCBuilder().anchorLineStart().weight(1, 0).resetPos().fillHorizontal();
		panel.setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Destination"));
		panel.add(m_target.getComponentPanel(), gbc.build());
		return panel;
	}

	private void toggleOpenFileAfterExecOption() {
		// cannot be headless as we'd not have a dialog in this case
		final boolean isRemote = XlsFormatterApplyNodeModel.isHeadlessOrRemote();
		final boolean categorySupported = XlsFormatterApplyNodeModel
				.categoryIsSupported(m_target.getSettingsModel().getLocation().getFSCategory());
		m_openOutputFile.getModel().setEnabled(!isRemote && categorySupported);
		if (isRemote) {
			m_openOutputFileLbl.setIcon(SharedIcons.INFO_BALLOON.get());
			m_openOutputFileLbl.setText("Not support in remote job view");
		} else if (!categorySupported) {
			m_openOutputFileLbl.setIcon(SharedIcons.INFO_BALLOON.get());
			m_openOutputFileLbl.setText("Not support by the selected file system");
		} else {
			m_openOutputFileLbl.setIcon(null);
			m_openOutputFileLbl.setText("");
		}
	}

	@Override
	protected void saveSettingsTo(final NodeSettingsWO settings) throws InvalidSettingsException {
		m_source.saveSettingsTo(settings);
		m_target.saveSettingsTo(settings);
		m_openOutputFile.saveSettingsTo(settings);
		m_preserveSourceNumberFormats.saveSettingsTo(settings);
	}

	@Override
	protected void loadSettingsFrom(final NodeSettingsRO settings, final PortObjectSpec[] specs)
			throws NotConfigurableException {
		m_source.loadSettingsFrom(settings, specs);
		m_target.loadSettingsFrom(settings, specs);
		m_openOutputFile.loadSettingsFrom(settings, specs);
		m_preserveSourceNumberFormats.loadSettingsFrom(settings, specs);
		toggleOpenFileAfterExecOption();
	}

	@Override
	public void onClose() {
		m_source.onClose();
		m_target.onClose();
		super.onClose();
	}

}
