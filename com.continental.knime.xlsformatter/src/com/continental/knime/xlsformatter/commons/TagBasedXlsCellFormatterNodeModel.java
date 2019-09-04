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

package com.continental.knime.xlsformatter.commons;

import java.util.List;
import org.apache.poi.ss.util.CellAddress;
import org.knime.core.node.NodeModel;
import org.knime.core.node.port.PortType;

/**
 * This class extends NodeModel in order to bring protected functionality to child classes
 * for shared warning message logic regarding tags. 
 */
public abstract class TagBasedXlsCellFormatterNodeModel extends NodeModel {

	protected TagBasedXlsCellFormatterNodeModel(PortType[] inPortTypes, PortType[] outPortTypes) {
		super(inPortTypes, outPortTypes);
	}
	
	protected void warnOnNoMatchingTags(List<CellAddress> matchingTags, String searchedTag) {
		if (matchingTags == null || matchingTags.size() == 0)
			setWarningMessage(getWarningMessage(searchedTag));
	}
	
	protected void warnOnNoMatchingTags(List<CellAddress> matchingTags, String searchedTag,
			WarningMessageContainer warningMessageContainer) {
		if (matchingTags == null || matchingTags.size() == 0)
			warningMessageContainer.addMessage(getWarningMessage(searchedTag));
	}
	
	public static String getWarningMessage(String searchedTag) {
		return "Tag \"" + searchedTag + "\" was not found in any cell of the provided control table.";
	}
}