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

import org.knime.core.node.ExecutionContext;

public class ProgressTools {
	
	/**
	 * Shows a node operation status message of the form "text 1 of 2 (50%) text" 
	 * @param exec The execution context to show the progress message in. Can be null, the method has no effect in this case.
	 * @param prefix The text to show before the figures.
	 * @param currentIteration The current iteration (typically starting at 1).
	 * @param total The total number of operations to calculate.
	 * @param postfix The text to show behind the figures.
	 */
	public static void showProgressText(final ExecutionContext exec, String prefix, int currentIteration, int total, String postfix) {
		
		if (exec == null)
			return;
		
		if (prefix != null && prefix.length() != 0 && !prefix.endsWith(" "))
			prefix += " ";
		else if (prefix == null)
			prefix = "";		
		
		if (postfix != null && postfix.length() != 0 && !postfix.startsWith(" "))
			postfix = " " + postfix;
		else if (postfix == null)
			postfix = "";
		
		exec.setProgress(prefix + currentIteration + " of " + total + " (" + Math.floor((double)currentIteration * 100 / total) + "%)" + postfix);
	}
}
