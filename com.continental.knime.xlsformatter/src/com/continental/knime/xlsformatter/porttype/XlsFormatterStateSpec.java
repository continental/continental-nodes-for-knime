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
	private static final long masterSerializationVersion = 3L;
	
	@Override
	public JComponent[] getViews() {
		
		return new JComponent[] { };
	}
	
	@Override
	public String toString() {
		return "XlsFormatterStateSpec";
	}
	
	public static XlsFormatterStateSpec getEmptySpec() {
		return new XlsFormatterStateSpec();
	}
	
	public XlsFormatterStateSpec getCopy() {
		XlsFormatterStateSpec ret = new XlsFormatterStateSpec();
		// if there are fields to this spec in the future, assign the member value to ret here
		return ret;
	}

	@Override
	public void writeExternal(ObjectOutput out) throws IOException {
		out.writeLong(masterSerializationVersion);
		out.writeBoolean(false); // due to version 1L
		
		/* Earliest old serialization version that this file can still be read/deserialized with.
		 * This concept provides a chance to selectively allow upward compatibility. (Downward compatibility shall always be
		 * supported.) When reading a file written by this method in an older code version, the above written masterSerializationVersion
		 * is first checked. Naively, the old code would throw an exception if a newer (i.e. higher) version is read (because
		 * the newer code could define a different byte stream even for the beginning of the file). However, newer code
		 * can write the earliest previous serialization version that is already capable to read this newer byte stream,
		 * despite not knowing about the additional content (i.e. typically since the changes in the versions since then
		 * are only appended at the end of the byte stream).
		 * 
		 * History:
		 * version 1, (selective upward compatibility not yet implemented)
		 * version 3, earliestSerializationVersionCapableOfReadingThis 1 (because new features are only added at the end of the byte stream)
		 */
		out.writeLong(1L); // WARNING: if unsure, set this one to masterSerializationVersion
	}

	@Override
	public void readExternal(ObjectInput in) throws IOException, ClassNotFoundException {
		long readSerializationVersion = in.readLong(); // the masterSerializationVersion of the code state that this object has originally been written with
		in.readBoolean(); // due to deprecated containsMergeInstruction in version 1L
		if (readSerializationVersion >= 3L) {
			long earliestSerializationVersionCapableOfReadingThis = in.readLong();
			if (masterSerializationVersion < readSerializationVersion && masterSerializationVersion < earliestSerializationVersionCapableOfReadingThis)
				throw new ClassNotFoundException("You are trying to read a XLS Formatting node configuration that has been written with a newer version of the extension than the one currently executing. Upward compatibility is not supported. Please update to the newest version.");
		}
	}
}
