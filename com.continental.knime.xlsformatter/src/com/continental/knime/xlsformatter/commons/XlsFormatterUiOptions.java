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

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import com.continental.knime.xlsformatter.porttype.XlsFormatterState;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState.FormattingFlag;

/**
 * Re-usable UI options (e.g. drop-down menus) 
 */
public class XlsFormatterUiOptions {

	private static List<String> _dropdownListOnOffUnmodified = null;
	
	public final static String UI_LABEL_SINGLE_TAG = "applies to tag (single tag only)";
	public final static String UI_LABEL_FONT = "(RGB, either in format #FF0000 or 255/0/0)";
	public final static String UI_TOOLTIP_SINGLE_TAG = "Select the tag from your control table for which the format should be changed.";
	public final static String UI_LABEL_CONTROLTABLESTYLE_KEY = "ControlTableStyle";
	public final static String UI_LABEL_CONTROLTABLESTYLE_GROUPTITLE = "Control Table Style";
	public final static String UI_LABEL_CONTROLTABLESTYLE_STANDARD = "standard tags";
	public final static String UI_ERROR_EMPTY_TAG = "Tag cannot be empty. Please set a freely chosen tag that matches an entry in the provided control table.";
	
	
	public static List<String> getDropdownListOnOffUnmodified() {
		if (_dropdownListOnOffUnmodified == null)
			_dropdownListOnOffUnmodified = getDropdownListFromEnum(XlsFormatterState.FormattingFlag.values());
		return _dropdownListOnOffUnmodified;
	}
	
	public static <T> List<String> getDropdownListFromEnum(T[] values) {
		return Arrays.stream(values).map(x -> x.toString().toLowerCase()).collect(Collectors.toList());
	}
	
	public static <T> String[] getDropdownArrayFromEnum(T[] values) {
		return Arrays.stream(values).map(x -> x.toString().toLowerCase()).toArray(String[]::new);
	}
	
	public static <T> T getEnumEntryFromString(T[] enumValues, String value) throws IllegalArgumentException {
		return Arrays.stream(enumValues).filter(v -> v.toString().equals(value)).findAny().orElseThrow(IllegalArgumentException::new);
	}
	
	public static XlsFormatterState.FormattingFlag getFormattingFlagFromDropdown(String value) throws Exception {
		return XlsFormatterState.FormattingFlag.valueOf(value.toUpperCase());
	}
	
	/**
	 * The reduced UI shows boolean, where false means unmodified (NOT off) in our internal formatting flag logic.
	 */
	public static XlsFormatterState.FormattingFlag getFormattingFlagFromBoolean(boolean value) throws Exception {
		return value ? FormattingFlag.ON : FormattingFlag.UNMODIFIED;
	}
}
