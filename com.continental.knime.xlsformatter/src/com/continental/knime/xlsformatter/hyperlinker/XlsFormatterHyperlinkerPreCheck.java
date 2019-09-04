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

package com.continental.knime.xlsformatter.hyperlinker;

import java.util.Map.Entry;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsFormatterHyperlinkerPreCheck  {

	/**
	 * Checks the passed hyperlinks for validity via the POI functionality that is also
	 * used in the Apply node's code. That way, warnings can be shown in the Hyperlink
	 * node already.
	 * An exception is thrown (with info about the causing cell and hyperlink) in case
	 * of an invalid contained link.
	 */
	public static void validateHyperlinks(Iterable<Entry<CellAddress, String>> addressToUrlMappings) {
		
		Workbook wb = null;
		try {
			wb = new XSSFWorkbook();
			CreationHelper createHelper = wb.getCreationHelper();
			XSSFHyperlink link = (XSSFHyperlink)createHelper.createHyperlink(HyperlinkType.URL);
			for (Entry<CellAddress, String> entry : addressToUrlMappings)
				try {
					link.setAddress(entry.getValue());
				} catch (Exception urlException) {
					throw new IllegalArgumentException("Hyperlink of cell " + entry.getKey().formatAsString() + " is not a valid URI: \"" + entry.getValue() + "\". See the node description for details.");
				}
		}
		finally {
			try {
				if (wb != null)
					wb.close();
			} catch (Exception ie) {}
		}
	}
}
