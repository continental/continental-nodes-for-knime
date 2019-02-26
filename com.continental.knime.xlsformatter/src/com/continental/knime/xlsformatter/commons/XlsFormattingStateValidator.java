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

/**
 * Validation for an intermediate XlsFormattingState checking whether it can still be implemented
 * according to the XLSX workbook specifications and limits as published in
 * https://support.office.com/en-us/article/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
 */
public class XlsFormattingStateValidator {

	public static final int MAX_CELL_STYLES_PER_WORKBOOK = 64000;
	public static final int MAX_FILL_STYLES_PER_WORKBOOK = 256;
	public static final int MAX_LINE_STYLES_PER_WORKBOOK = 256;
	public static final int MAX_FONTS_PER_WORKBOOK = 512;
	public static final int MAX_NUMBER_FORMATS_PER_WORKBOOK = 200;
	public static final int MAX_ROW_HEIGHT_IN_POINTS = 409;
	public static final int MAX_HYPERLINKS_PER_WORKBOOK = 66530;
	public static final int MAX_FORMULA_CHARACTERS = 8192;
}
