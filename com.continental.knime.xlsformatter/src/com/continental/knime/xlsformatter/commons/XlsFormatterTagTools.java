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

import org.apache.commons.lang3.StringUtils;

public class XlsFormatterTagTools {

	public final static String INVALID_TAGLIST_CHARACTERS = ";*/|+&?!";
	public final static String INVALID_TAG_CHARACTERS = INVALID_TAGLIST_CHARACTERS + ",";
	
	/**
	 * Parse a tag field by interpreting it as a comma-separated list.
	 */
	public static List<String> parseCommaSeparatedTagList(String tagString) {
		String[] tagList = tagString.split(","); 
		for (int i = 0; i < tagList.length; i++)
			tagList[i] = tagList[i].trim();
		return Arrays.asList(tagList);
	}	
	
	/**
	 * Check whether a single search tag is present in a comma-separated list of tags.
	 */
	public static boolean doesTagMatch(String cellTags, String searchTag) {
		return parseCommaSeparatedTagList(cellTags).stream()
				.filter(s -> s.trim().length() >= 1 && s.equals(searchTag.trim()))
				.findFirst().isPresent();
	}
	
	/**
	 * Checks whether a single tag is valid, i.e. doesn't have any forbidden characters, esp. comma.
	 */
	public static boolean isValidSingleTag(String tag) {
		return tag.trim().length() >= 1 && !StringUtils.containsAny(tag, INVALID_TAG_CHARACTERS);
	}
	
	/**
	 * Checks whether a comma-separated tag list is valid, i.e. doesn't have any forbidden characters
	 */
	public static boolean isValidTagList(String commaSeparatedTagList) {
		return commaSeparatedTagList.trim().length() >= 1 && !StringUtils.containsAny(commaSeparatedTagList, INVALID_TAGLIST_CHARACTERS);
	}
}
