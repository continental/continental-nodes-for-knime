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

package com.continental.knime.xlsformatter.xlscontroltablemerger;

import java.util.Optional;

import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.ConfigurableNodeFactory;
import org.knime.core.node.NodeDialogPane;
import org.knime.core.node.NodeView;
import org.knime.core.node.context.NodeCreationConfiguration;
import org.knime.core.node.port.PortType;

public class XlsControlTableMergerNodeFactory extends ConfigurableNodeFactory<XlsControlTableMergerNodeModel> {

	/**
	 * {@inheritDoc}
	 */
	@Override
	public XlsControlTableMergerNodeModel createNodeModel() {
		return new XlsControlTableMergerNodeModel();
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
	public NodeView<XlsControlTableMergerNodeModel> createNodeView(final int viewIndex,
			final XlsControlTableMergerNodeModel nodeModel) {
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
	protected Optional<PortsConfigurationBuilder> createPortsConfigBuilder() {
		PortsConfigurationBuilder builder = new PortsConfigurationBuilder();
    builder.addExtendableInputPortGroup(
    		"input",
    		new PortType[]{ BufferedDataTable.TYPE, BufferedDataTable.TYPE},
    		BufferedDataTable.TYPE);
    builder.addFixedOutputPortGroup("Merged XLS Control Table", BufferedDataTable.TYPE);
    return Optional.of(builder);
	}

	@Override
	protected XlsControlTableMergerNodeModel createNodeModel(NodeCreationConfiguration creationConfig) {
    return new XlsControlTableMergerNodeModel(creationConfig.getPortConfig().get());
	}

	@Override
	protected NodeDialogPane createNodeDialogPane(NodeCreationConfiguration creationConfig) {
		return new XlsControlTableMergerNodeDialog();
	}
}
