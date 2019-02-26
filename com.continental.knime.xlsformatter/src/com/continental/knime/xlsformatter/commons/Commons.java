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

public class Commons {

	/**
	 * Parses an integer and returns null on mismatch.
	 */
	public static Integer parseIntSilently(String value) {
		try {
			return new Integer(Integer.parseInt(value));
		}
		catch (NumberFormatException ne) {
			return null;
		}
	}
	
	/**
	 * Parses an integer and returns an IllegalArgumentException on mismatch.
	 */
	public static int parseInt(String value) throws IllegalArgumentException {
		try {
			return new Integer(Integer.parseInt(value));
		}
		catch (NumberFormatException ne) {
			throw new IllegalArgumentException("Expected an integer value (i.e. a whole number), but saw \"" + value + "\".");
		}
	}
	
	/**
	 * Modes of how to add two tables.
	 */
	public static enum Modes {
  	APPEND,
  	OVERWRITE;
  	
  	@Override
		public String toString() {
			switch (this) {
			case APPEND:
				return "append";
			case OVERWRITE:
				return "overwrite";
			default:
				throw new IllegalArgumentException();
			}
		}
		
		public static Modes getFromString(String value) {
	    return XlsFormatterUiOptions.getEnumEntryFromString(Modes.values(), value);
		}	
  }
}
