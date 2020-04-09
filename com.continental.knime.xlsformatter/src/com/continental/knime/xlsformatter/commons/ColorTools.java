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

import java.awt.Color;

import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * Static methods helping to parse width / height instructions
 */
public class ColorTools {

	/**
	 * Converts a hex color string such as #FF0000 to our internal representation of 255/0/0.
	 * Returns null on failure.
	 */
	private static Color hexColorToXlsfColor(String hexColorCode) throws IllegalArgumentException {
		try {
			return Color.decode(hexColorCode);
		} catch (NumberFormatException ne) {
			throw new IllegalArgumentException("Could not interpret RGB color code " + hexColorCode + ", expected something like #FF0000 for red.");
		}
	}
	
	/**
	 * Checks whether the provided argument is a color code of form #FF0000 or 255/0/0 and converts them to a color
	 */
	public static Color anyColorCodeToXlsfColor(String value) throws IllegalArgumentException {
		if (value.substring(0, 1).equals("#"))
			return hexColorToXlsfColor(value);
		String[] components = value.split("/");
		try {
		if (components.length != 3)
			throw new Exception();
			return new Color(Integer.parseInt(components[0]),
				Integer.parseInt(components[1]),
				Integer.parseInt(components[2]));
		} catch (Exception e) {
			throw new IllegalArgumentException("Could not interpret color code \"" + value + "\". Valid RGB formats for e.g. red would be #FF0000 or 255/0/0.");
		}
	}
	
	/**
	 * Creates a POI color from a java.awt.Color, which is used as Xls Formatter State internal color representation.
	 */
	public static XSSFColor getPoiColor(Color color) {
		return new XSSFColor(color); 		
	}
	
	/**
	 * Outputs a Color object in R/G/B form, respective A/R/G/B if alpha is not equal 255 
	 */
	public static String colorToXlsfColorString(Color value) {
		return (value.getAlpha() != 255 ? value.getAlpha() + "/" : "") + value.getRed() + "/" + value.getGreen() + "/" + value.getBlue();
	}
	public static String colorToXlsfColorString(XSSFColor value) {
		if (value == null)
			return null;
		byte[] rgbArray = value.getRGB(); // byte is signed in Java
		if (rgbArray == null || rgbArray.length != 3)
			return null;
		return (rgbArray[0] & 0xFF) + "/" + (rgbArray[1] & 0xFF) + "/" + (rgbArray[2] & 0xFF);
	}
}
