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

package com.continental.knime.xlsformatter.portmerger2;

import org.knime.core.node.NodeDialogPane;

import java.util.Optional;

import org.knime.core.node.ConfigurableNodeFactory;
import org.knime.core.node.context.NodeCreationConfiguration;
import org.knime.core.node.port.PortType;

import com.continental.knime.xlsformatter.porttype.XlsFormatterState;

import org.knime.core.node.NodeView;

public class XlsPortMergerNodeFactory extends ConfigurableNodeFactory<XlsPortMergerNodeModel> {

	/**
	 * {@inheritDoc}
	 */
	@Override
	public XlsPortMergerNodeModel createNodeModel() {
		return new XlsPortMergerNodeModel();
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public int getNrNodeViews() {
		return 0;
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public NodeView<XlsPortMergerNodeModel> createNodeView(final int viewIndex,
			final XlsPortMergerNodeModel nodeModel) {
		return null;
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	public boolean hasDialog() {
		return true;
	}

	@Override
  protected XlsPortMergerNodeModel createNodeModel(final NodeCreationConfiguration creationConfig) {
      return new XlsPortMergerNodeModel(creationConfig.getPortConfig().get());
  }

	@Override
	protected Optional<PortsConfigurationBuilder> createPortsConfigBuilder() {
    PortsConfigurationBuilder builder = new PortsConfigurationBuilder();
    builder.addExtendableInputPortGroup(
    		"input",
    		new PortType[]{ XlsFormatterState.TYPE, XlsFormatterState.TYPE},
    		XlsFormatterState.TYPE);
    builder.addFixedOutputPortGroup("Merged XLS Formatter State", XlsFormatterState.TYPE);
    return Optional.of(builder);
	}

	@Override
	protected NodeDialogPane createNodeDialogPane(NodeCreationConfiguration creationConfig) {
		return new XlsPortMergerNodeDialog();
	}
}
