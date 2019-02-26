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

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.defaultnodesettings.SettingsModelString;

/**
 * Re-usable UI options (e.g. drop-down menus) 
 */
public class UiValidation {

	public final static String UI_ERROR_EMPTY_TAG = "Tag cannot be empty. Please set a freely chosen tag that matches an entry in the provided control table.";
	public final static String UI_ERROR_INVALID_TAG = "Tag contains invalid characters (" + XlsFormatterTagTools.INVALID_TAG_CHARACTERS + ").";
	
	
	public static void validateTagField(SettingsModelString tag, String additionalEmptyErrorMessage) throws InvalidSettingsException {
		
		if (tag.getStringValue().trim().equals(""))
			throw new InvalidSettingsException(UI_ERROR_EMPTY_TAG + (additionalEmptyErrorMessage == null ? "" : "\n" + additionalEmptyErrorMessage));
		
		if (!XlsFormatterTagTools.isValidSingleTag(tag.getStringValue()))
			throw new InvalidSettingsException(UI_ERROR_INVALID_TAG);
	}
	
	public static void validateTagField(SettingsModelString tag) throws InvalidSettingsException {
		validateTagField(tag, null);
	}
}
