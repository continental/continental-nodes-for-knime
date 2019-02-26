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

package com.continental.knime.utility.networkcomponentsplitter;

import java.util.AbstractMap;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Set;

import javax.naming.directory.InvalidAttributesException;

import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataType;
import org.knime.core.data.MissingCell;
import org.knime.core.data.RowKey;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.IntCell;
import org.knime.core.data.def.StringCell;
import org.knime.core.data.sort.BufferedDataTableSorter;
import org.knime.core.node.BufferedDataContainer;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.NodeLogger;

public class NetworkComponentSplitterNodeLogic {
	
	public static BufferedDataTable[] execute(final BufferedDataTable[] inData,
			final String inputColumnName1, final String inputColumnName2, final boolean missingValueAsOwnNode,
			final String outputNodeColumnName, final String outputClusterColumnName, 
			final ExecutionContext exec, final int IN_PORT, final NodeLogger logger) throws Exception {
		
  	// Your custom variables:
		HashMap<String, Integer> nodeToClusterMap = new HashMap<String, Integer>();
		int nextFreeClusterId = 0;
		HashSet<AbstractMap.SimpleEntry<Integer, Integer>> clusterConnections = new HashSet<AbstractMap.SimpleEntry<Integer, Integer>>();
		
		/*
		 * Idea of this algorithm:
		 * Part 1: Process the input edge table row by row (only one pass of this table ever read). While traversing this table,
		 * fill a dictionary mapping every node to a cluster id. At any new edge, check if one of the nodes has been seen before
		 * and assign the same cluster to it. Only if both nodes of the edge are new, assign the next free cluster id.
		 * Part 2: At the end of the above process, we need to make sure that the result is consistent. Some post-processing is
		 * obviously necessary in cases where e.g. two big clusters where filled during the process and in one of the last rows
		 * we find a link between these clusters in the data (and hence by definition making all nodes of both assumed clusters
		 * in fact belonging to the very same cluster). The algorithmic strategy for this job is to keep track via two dictionaries.
		 * In the "smallestRelatedClusterMap" we map every initial cluster id to the smallest update that we know of. This means that
		 * these clusters will end up as the same cluster in the final result. Any time we change something in this dictionary, we
		 * would need to traverse through all other dictionary entries and check whether the changed value also occurs there. This is
		 * not wise in terms of computational complexity. Hence we track another dictionary "clusterGroupMap" that can be understood
		 * a reverse direction of the smallestRelatedClusterMap. That means its keys are the values of smallestRelatedClusterMap, and
		 * its values are lists of all of smallestRelatedClusterMap's keys mapping to these. So e.g. 1->1, 2->1, 3->3 in the forward
		 * map would yield 1->{1,2}, 3->{3} in the reverse map. That way, any changes to one of the maps can directly be applied to
		 * all subsequent entries.
		 * Part 3: As a cosmetic next step, we want to make sure that the final cluster ids are 1, 2, 3, ... without gaps in
		 * numbering. Since the above logic has introduced these gaps, we need to re-assign ids.
		 */
		
		logger.info("Current Status of Cluster Node: 0%");
		
		BufferedDataTable inputTable = inData[IN_PORT];
		BufferedDataContainer outBuffer = exec.createDataContainer(
				new DataTableSpec(
						DataTableSpec.createColumnSpecs(
								new String[] { outputNodeColumnName, outputClusterColumnName },
								new DataType[] { StringCell.TYPE, IntCell.TYPE })));
		
		long rowCount = inputTable.size();
		java.util.List<String> tableColumnNames = Arrays.asList(inputTable.getSpec().getColumnNames());
		int colIndexNode1 = tableColumnNames.indexOf(inputColumnName1);
		int colIndexNode2 = tableColumnNames.indexOf(inputColumnName2);
		
		if (colIndexNode1 == -1)
			throw new InvalidAttributesException("Column " + inputColumnName1 + " not found.");
		if (colIndexNode2 == -1)
			throw new InvalidAttributesException("Column " + inputColumnName2 + " not found.");
		
		// read in data table and fill HashMaps in this first pass right away
		long currentRow = 0;
		for (DataRow row : inputTable) {
			
			exec.checkCanceled();
			exec.setProgress(currentRow / rowCount * 0.5d,
					"Read In Table: " + Math.round(currentRow/rowCount*100d) + "%");
			
			String key1 =
					row.getCell(colIndexNode1).isMissing() || row.getCell(colIndexNode1).getType() != StringCell.TYPE ? null
							: row.getCell(colIndexNode1).toString();
			String key2 =
					row.getCell(colIndexNode2).isMissing() || row.getCell(colIndexNode2).getType() != StringCell.TYPE ? null
							: row.getCell(colIndexNode2).toString();
			
			// Resolve missing value in case the respective UI setting disallows missing/null as valid node
			if (!missingValueAsOwnNode) {
				if (key1 == null && key2 == null)
					continue; // skip the entire data row
				if (key1 == null)
					key1 = key2;
				if (key2 == null)
					key2 = key1;
			}
			
			// core logic of first pass through the data: initially fill the hashmaps
			Integer cluster1 = nodeToClusterMap.containsKey(key1) ? nodeToClusterMap.get(key1) : null;
			Integer cluster2 = nodeToClusterMap.containsKey(key2) ? nodeToClusterMap.get(key2) : null;
	
			if (cluster1 != null) { // key1 is known
				if (cluster2 == null) // is key2 unknown?
					nodeToClusterMap.put(key2, cluster1);
				else if ((int)cluster1 != (int)cluster2)
					clusterConnections.add(new AbstractMap.SimpleEntry<Integer, Integer>(Math.min((int)cluster1, (int)cluster2), Math.max((int)cluster1, (int)cluster2)));
			}
			else if (cluster2 != null) { // key2 is known, (key1 is unknown)
				nodeToClusterMap.put(key1, cluster2); // because key1 is definitely unknown (otherwise we wouldn't be in this else-if)
			}
			else { // both nodes of this edge have not been seen before
				nodeToClusterMap.put(key1, nextFreeClusterId);
				nodeToClusterMap.put(key2, nextFreeClusterId);
				nextFreeClusterId++;
			}
			
			currentRow++;
		}
  		
		
		// consolidate the information now that the entire table has been processed once
		logger.info("Current Status of Cluster Node: 50%");
		HashMap<Integer, Integer> smallestRelatedClusterMap = new HashMap<Integer, Integer>(); // per every previously assigned cluster, keep track of the smallest linked cluster currently known
		HashMap<Integer, HashSet<Integer>> clusterGroupMap = new HashMap<Integer, HashSet<Integer>>(); // reverse look-up table (in above logic: which value maps to which keys)
		
		// initially fill new data structures
		long progress = 0;
		long total = nextFreeClusterId;  			
		for (int i = 0; i < nextFreeClusterId; i++) {
			exec.checkCanceled();
			exec.setProgress((double) 0.5 + progress/total*0.075,
					"Calculate ClusterIDs for Nodes " + Math.round((double) progress++/total*100*0.3/0.4) + "%");
			
			smallestRelatedClusterMap.put(i, i); // add self-reference for all clusters not appearing in a connection
			clusterGroupMap.put(i, new HashSet<Integer>());
			clusterGroupMap.get(i).add(i);
		}
		
		// iterate through the known connections between clusters
		progress = 0;
		total = clusterConnections.size();
		for (AbstractMap.SimpleEntry<Integer, Integer> pair : clusterConnections) {
			exec.checkCanceled();
			exec.setProgress((double) 0.575 + progress/total*0.025,
					"Calculate ClusterIDs for Nodes " + Math.round((double) 75+progress++/total*100*0.1/0.4) + "%");
			Integer pairSmaller = pair.getKey(); // by the above adding strategy, key is always smaller than value in these connections
			Integer pairLarger = pair.getValue();
			Integer valuePairS = smallestRelatedClusterMap.get(pairSmaller);
			Integer valuePairL = smallestRelatedClusterMap.get(pairLarger);
			Integer valueSmaller = Math.min((int)valuePairS, (int)valuePairL);
			Integer valueLarger = Math.max((int)valuePairS, (int)valuePairL);
		
			// decide whether action needs to be taken or whether these two clusters where already previously merged (i.e. map to the same cluster already)
			if ((int)valueSmaller != (int)valueLarger) {
				
				// use the "backward facing" map to determine which updates are needed in the "forward facing" map:
				for (Integer i : clusterGroupMap.get(valueLarger)) {
					exec.checkCanceled();
					smallestRelatedClusterMap.put(i, valueSmaller);
				}
		
				// update the "backward facing" map:
				clusterGroupMap.get(valueSmaller).addAll(clusterGroupMap.get(valueLarger));
				clusterGroupMap.remove(valueLarger);
			}
		}
		clusterConnections = null;
		clusterGroupMap = null;
		
		// re-label the clusters in order to be consecutive again
		HashMap<Integer, Integer> finalClusterIds = new HashMap<Integer, Integer>(); // intermediate cluster ids to final ids
		nextFreeClusterId = 0;
		Set<Integer> keys = smallestRelatedClusterMap.keySet();
		
		progress = 0;
		total = keys.size();
		for (Integer key : keys) {
			exec.checkCanceled();
			exec.setProgress(0.6d + progress / total * 0.05,
					"Generate final ClusterIDs: " + Math.round((double) progress++ / total * 100 * 0.1 / 0.4) + "%");
			Integer right = smallestRelatedClusterMap.get(key);
			if (!finalClusterIds.containsKey(right))
				finalClusterIds.put(right, nextFreeClusterId++);
			smallestRelatedClusterMap.put(key, finalClusterIds.get(right)); // directly assign the final ID
		}
		finalClusterIds = null;
		
		
		// assign nodes to clusters, i.e. create unsorted output table
		long i = 0;
		total = nodeToClusterMap.size();
		for (String key : nodeToClusterMap.keySet()) {
			exec.checkCanceled();
			exec.setProgress(0.8d + i / total * 0.1,
					"Write unsorted output Table: " + (i / total * 100d) + "%");
			
			DataCell[] cells = new DataCell[2];
			cells[0] = key == null ? new MissingCell("") : new StringCell(key);
			cells[1] = new IntCell(smallestRelatedClusterMap.get(nodeToClusterMap.get(key)) + 1); // + 1 to have a 1-based instead of 0-based cluster id list
			DataRow rowOut = new DefaultRow(RowKey.createRowKey(i), cells);
			outBuffer.addRowToTable(rowOut);
			i++;
		}
  	
  	outBuffer.close();
  	exec.setProgress(0.90d, "Sort Output Table");
  	
  	
  	// sort output table
  	BufferedDataTable[] outTables = new BufferedDataTable[1];
  	outTables[0] = outBuffer.getTable();
  	Collection<String> sortColumnOrder = Arrays.asList(outputClusterColumnName, outputNodeColumnName);
  	outTables[0] = (new BufferedDataTableSorter(outTables[0], sortColumnOrder, new boolean[] {true, true}, true))
  			.sort(exec);
  	
  	// assign row keys
  	i = 0;
  	outBuffer = exec.createDataContainer(outTables[0].getSpec());
  	for (DataRow row : outTables[0]) {
  		exec.checkCanceled();
  		DataRow rowOut = new DefaultRow(RowKey.createRowKey(i++), row);
  		outBuffer.addRowToTable(rowOut);
  	}
  	outBuffer.close();
  	outTables[0] = outBuffer.getTable();
  	exec.setProgress(100d);
  	return outTables;
	}
}
