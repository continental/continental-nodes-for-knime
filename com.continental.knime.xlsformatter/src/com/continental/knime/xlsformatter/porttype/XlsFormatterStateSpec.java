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

package com.continental.knime.xlsformatter.porttype;

import java.io.Externalizable;
import java.io.IOException;
import java.io.ObjectInput;
import java.io.ObjectOutput;

import javax.swing.JComponent;

import org.knime.core.node.port.PortObjectSpec;

public class XlsFormatterStateSpec implements PortObjectSpec, Externalizable {
	
	private static final long serialVersionUID = 1L;
	
	/**
	 * Defines whether the chain of XLS Formatter instruction so far has seen any cell merge instruction.
	 * Used to deny further adding non-cell merge operations as cell merge needs to come last in the chain.
	 */
	private boolean _containsMergeInstruction = false;
	
	@Override
	public JComponent[] getViews() {
		
		return new JComponent[] { };
	}
	
	public boolean getContainsMergeInstruction() { return _containsMergeInstruction; }
	public void setContainsMergeInstruction(boolean value) { _containsMergeInstruction = value; }
	
	@Override
	public String toString() {
		return "XlsFormatterStateSpec";
	}
	
	public static XlsFormatterStateSpec getEmptySpec() {
		return new XlsFormatterStateSpec();
	}
	
	public XlsFormatterStateSpec getCopy() {
		XlsFormatterStateSpec ret = new XlsFormatterStateSpec();
		ret._containsMergeInstruction = _containsMergeInstruction;
		return ret;
	}

	@Override
	public void writeExternal(ObjectOutput out) throws IOException {
		out.writeLong(serialVersionUID);
		out.writeBoolean(_containsMergeInstruction);
	}

	@Override
	public void readExternal(ObjectInput in) throws IOException, ClassNotFoundException {
		in.readLong(); // long readSerialVersionUid
		_containsMergeInstruction = in.readBoolean();
	}
}
