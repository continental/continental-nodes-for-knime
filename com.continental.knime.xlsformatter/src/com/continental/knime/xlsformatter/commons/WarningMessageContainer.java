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

import java.util.HashSet;
import java.util.Set;

/**
 * Class that enables passing out a side parameter regarding a warning message despite successful node execution
 * to the calling NodeModel.execute() function. 
 */
public class WarningMessageContainer {
	
	private String m_message = null;
	
	private Set<String> m_containedMessages = new HashSet<String>();
	
	public void addMessage(String message) {
		if (m_message == null) {
			m_message = message;
			m_containedMessages.add(message);
		}
		else if (!m_containedMessages.contains(message)) {
			m_message = m_message + " " + message;
			m_containedMessages.add(message);
		}
	}
	public String getMessage() {
		return m_message;
	}
	public boolean hasMessage() {
		return m_message != null;
	}
}
