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

package com.continental.knime.xlsformatter.cellbackgroundcolorizer;

import org.knime.core.node.NodeView;

public class XlsFormatterCellBackgroundColorizerNodeView extends NodeView<XlsFormatterCellBackgroundColorizerNodeModel> {

	/**
	 * Creates a new view.
	 * 
	 * @param nodeModel The model (class: {@link XlsFormatterCellFormatterNodeModel})
	 */
	protected XlsFormatterCellBackgroundColorizerNodeView(final XlsFormatterCellBackgroundColorizerNodeModel nodeModel) {
		super(nodeModel);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void modelChanged() {

		XlsFormatterCellBackgroundColorizerNodeModel nodeModel = 
				(XlsFormatterCellBackgroundColorizerNodeModel)getNodeModel();
		assert nodeModel != null;

		// be aware of a possibly not executed nodeModel! The data you retrieve
		// from your nodemodel could be null, emtpy, or invalid in any kind.
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void onClose() {
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
	protected void onOpen() {
	}
}