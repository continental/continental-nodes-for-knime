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

import java.util.Optional;

import org.knime.core.node.ConfigurableNodeFactory;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeView;
import org.knime.core.node.context.NodeCreationConfiguration;
import org.knime.core.node.context.ports.PortsConfiguration;
import org.knime.filehandling.core.port.FileSystemPortObject;

import com.continental.knime.xlsformatter.porttype.XlsFormatterState;

public final class XlsFormatterApplyNodeFactory extends ConfigurableNodeFactory<XlsFormatterApplyNodeModel> {

	/** The name of the optional source connection input port group. */
	private static final String CONNECTION_SOURCE_PORT_GRP_NAME = "Source File System Connection";

	/** The name of the optional destination connection input port group. */
	private static final String CONNECTION_DESTINATION_PORT_GRP_NAME = "Destination File System Connection";

	/** The name of the xls formatter input port group. */
	private static final String FORMATTER_STATE_GRP_NAME = "XLS Formatter";

	@Override
	protected Optional<PortsConfigurationBuilder> createPortsConfigBuilder() {
		final PortsConfigurationBuilder b = new PortsConfigurationBuilder();
		b.addOptionalInputPortGroup(CONNECTION_SOURCE_PORT_GRP_NAME, FileSystemPortObject.TYPE);
		b.addFixedInputPortGroup(FORMATTER_STATE_GRP_NAME, XlsFormatterState.TYPE);
		b.addOptionalInputPortGroup(CONNECTION_DESTINATION_PORT_GRP_NAME, FileSystemPortObject.TYPE);
		return Optional.of(b);
	}

	@Override
	protected XlsFormatterApplyNodeModel createNodeModel(NodeCreationConfiguration creationConfig) {
		final PortsConfiguration portsCfg = getPortsCfg(creationConfig);
		return new XlsFormatterApplyNodeModel(portsCfg, createSettings(portsCfg), getFormatterIdx(portsCfg));
	}

	@Override
	protected NodeDialogPane createNodeDialogPane(NodeCreationConfiguration creationConfig) {
		final PortsConfiguration portsCfg = getPortsCfg(creationConfig);
		return new XlsFormatterApplyNodeDialog(createSettings(portsCfg));
	}

	@Override
	protected int getNrNodeViews() {
		return 0;
	}

	@Override
	public NodeView<XlsFormatterApplyNodeModel> createNodeView(int viewIndex, XlsFormatterApplyNodeModel nodeModel) {
		return null;
	}

	@Override
	protected boolean hasDialog() {
		return true;
	}

	private static PortsConfiguration getPortsCfg(final NodeCreationConfiguration creationCfg) {
		return creationCfg.getPortConfig().orElseThrow(IllegalStateException::new);
	}

	private static XlsFormatterApplySettings createSettings(final PortsConfiguration portsCfg) {
		return new XlsFormatterApplySettings(portsCfg, CONNECTION_SOURCE_PORT_GRP_NAME,
				CONNECTION_DESTINATION_PORT_GRP_NAME);
	}

	private static int getFormatterIdx(final PortsConfiguration portsCfg) {
		return portsCfg.getInputPortLocation().get(FORMATTER_STATE_GRP_NAME)[0];
	}
}
