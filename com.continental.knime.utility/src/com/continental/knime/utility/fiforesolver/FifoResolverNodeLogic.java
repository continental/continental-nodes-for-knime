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

package com.continental.knime.utility.fiforesolver;

import java.util.AbstractMap;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.Deque;  // double ended queue, better than Queue/Stack

import javax.naming.directory.InvalidAttributesException;

import org.knime.core.data.DataCell;
import org.knime.core.data.DataRow;
import org.knime.core.data.DataTableSpec;
import org.knime.core.data.DataType;
import org.knime.core.data.DoubleValue;
import org.knime.core.data.MissingCell;
import org.knime.core.data.RowKey;
import org.knime.core.data.def.DefaultRow;
import org.knime.core.data.def.DoubleCell;
import org.knime.core.data.def.IntCell;
import org.knime.core.data.def.LongCell;
import org.knime.core.data.def.StringCell;
import org.knime.core.node.BufferedDataContainer;
import org.knime.core.node.BufferedDataTable;
import org.knime.core.node.ExecutionContext;
import org.knime.core.node.ExecutionMonitor;
import org.knime.core.node.NodeLogger;

public class FifoResolverNodeLogic {
	
	public static BufferedDataTable[] execute(final BufferedDataTable[] inData, final boolean isModeFifo,
			final String inputColumnNameGroup, final String inputColumnNameQty, final boolean failAtInconsistency,
			final ExecutionContext exec, final int IN_PORT, final NodeLogger logger) throws Exception {
		
		// qm (queue map): <group name, <row id IN, remaining quantity>>
		HashMap<String, FifoLifoStore> qm = new HashMap<String, FifoLifoStore>();
		ExecutionMonitor subExecMonitor;
		
		BufferedDataTable inputTable = inData[IN_PORT];
		long rowCount = inputTable.size();
		java.util.List<String> tableColumnNames = Arrays.asList(inputTable.getSpec().getColumnNames());
		int colIndexGroup = tableColumnNames.indexOf(inputColumnNameGroup);
		int colIndexQty = tableColumnNames.indexOf(inputColumnNameQty);
		
		if (colIndexGroup == -1)
			throw new InvalidAttributesException("Column " + inputColumnNameGroup + " not found.");
		if (colIndexQty == -1)
			throw new InvalidAttributesException("Column " + inputColumnNameQty + " not found.");
		
		
		DataType qtyNumberType = inputTable.getSpec().getColumnSpec(inputColumnNameQty).getType();
		String qtyNumberTypeString = qtyNumberType.getName();
		BufferedDataContainer outBuffer = exec.createDataContainer(
				new DataTableSpec(
						DataTableSpec.createColumnSpecs(
								new String[] { inputColumnNameGroup, "RowId_IN", "RowId_OUT", inputColumnNameQty },
								new DataType[] { StringCell.TYPE, StringCell.TYPE, StringCell.TYPE, qtyNumberType })));
		
		// main loop of processing the input table (and where possible already populating the output table)
		long currentRow = 0;
		subExecMonitor = exec.createSubProgress(.8d);
		for (DataRow row : inputTable) {
			subExecMonitor.checkCanceled();
			subExecMonitor.setProgress((double)currentRow / rowCount,
					"process input table: " + Math.round(currentRow * 100d / rowCount) + "%");
			
			// read data from row:
			String group =
					row.getCell(colIndexGroup).isMissing() || row.getCell(colIndexGroup).getType() != StringCell.TYPE ? null
							: row.getCell(colIndexGroup).toString();
			Double qty = null;
			if (!(row.getCell(colIndexQty).isMissing() || !row.getCell(colIndexQty).getType().isCompatible(DoubleValue.class))) {
				if (qtyNumberTypeString == "Number (long)")
					qty = ((LongCell)row.getCell(colIndexQty)).getDoubleValue();
				else if (qtyNumberTypeString == "Number (integer)")
					qty = ((IntCell)row.getCell(colIndexQty)).getDoubleValue();
				else
					qty = ((DoubleCell)row.getCell(colIndexQty)).getDoubleValue();
			}
			
			String rowId = row.getKey().getString();
			
			// missing value handling:
			if (qty == null)
				continue;
			
			// main FIFO/LIFO handling:
			if (!qm.containsKey(group))
				qm.put(group, new FifoLifoStore(isModeFifo));
			FifoLifoStore q = qm.get(group); // get this group's queue
			
			if (qty >= 0)  // IN: just add to queue!
				q.add(new AbstractMap.SimpleEntry<String, Double>(rowId, qty));
			
			else {  // OUT: handle as new out row(s)
				AbstractMap.SimpleEntry<String, Double> p = null;
				
				while (qty < 0) {  // iteratively reduce the OUT quantity based on previous IN in queue
					Double v = null;
					p = q.peek();
					if (p == null) {  // inconsistency, because we still have OUT qty, but nothing in queue anymore
						if (failAtInconsistency)
							throw new Exception("Row ID \"" + rowId + "\" indicates an outflow beyond previous inflow for group \"" + group + "\". The node setting is configured to fail at such queuing inconsistencies.");
						else {  // add as output table row with missing input row-id
							DataCell[] cells = new DataCell[4];
							cells[0] = group == null ? new MissingCell("") : new StringCell(group);
							cells[1] = new MissingCell("");
							cells[2] = new StringCell(rowId);
							cells[3] = generateCell(-qty, qtyNumberTypeString);
							outBuffer.addRowToTable(new DefaultRow(new RowKey("PHANTOM_" + rowId), cells));
							qty = 0.0;
						}
					} else {   // we still have something queued to deduct this OUT qty from				
						v = p.getValue();
						Double outputQty = null;
						String inputRowId = p.getKey();
						if (v <= -qty) {   // the entire IN queue element needs to be consumed
							q.poll();
							outputQty = v;
							qty += v;   // qty is negative, v is positive
						} else {   // the IN queue element is bigger than the OUT qty, so reduce but leave it
							p.setValue(v + qty);
							outputQty = -qty;
							qty = 0.0;
						}
						DataCell[] cells = new DataCell[4];
						cells[0] = group == null ? new MissingCell("") : new StringCell(group);
						cells[1] = new StringCell(inputRowId);
						cells[2] = new StringCell(rowId);
						cells[3] = generateCell(outputQty, qtyNumberTypeString);
						outBuffer.addRowToTable(new DefaultRow(new RowKey(inputRowId + "_" + rowId), cells));
					}
				}
			}
		}
		
		// handle end inventory (still enqueued at end, meaning not matched with an OUT):
		subExecMonitor = exec.createSubProgress(.2d);
		long i = 0;
		for (String group : qm.keySet()) {
			subExecMonitor.checkCanceled();
			subExecMonitor.setProgress((double)i / (i+1),
					"post-process remainder: " + Math.round(i * 100d / (i+1)) + "%");
			FifoLifoStore q = qm.get(group); // get this group's queue
			AbstractMap.SimpleEntry<String, Double> entry = null;
			while ((entry = q.poll()) != null) {
				DataCell[] cells = new DataCell[4];
				cells[0] = group == null ? new MissingCell("") : new StringCell(group);
				cells[1] = new StringCell(entry.getKey());
				cells[2] = new MissingCell("");
				cells[3] = generateCell(entry.getValue(), qtyNumberTypeString);
				outBuffer.addRowToTable(new DefaultRow(new RowKey(entry.getKey() + " (residual)"), cells));
			}
		}
		
		// finish up:
		outBuffer.close();
		BufferedDataTable[] outTables = new BufferedDataTable[1];
	  	outTables[0] = outBuffer.getTable();
	  	return outTables;
	}
	
	/*
	 * Custom queue class handling the difference between FIFO and LIFO modes merely by setting
	 * the constructor boolean indicator.
	 * Implemented via double ended queue, where
	 *   - FIFO adds at the end and removes at the front of the queue
	 *   - LIFO adds and removes at the front of the queue
	 */
	private static class FifoLifoStore {
		private Deque<AbstractMap.SimpleEntry<String, Double>> store = new LinkedList<AbstractMap.SimpleEntry<String, Double>>();
		private boolean isModeFifo;
		
		public FifoLifoStore(boolean isModeFifo) {
			this.isModeFifo = isModeFifo;
		}
		
		public AbstractMap.SimpleEntry<String, Double> poll() {
			return store.pollFirst();
		}
		
		public AbstractMap.SimpleEntry<String, Double> peek() {
			return store.peekFirst();
		}
		
		public void add(AbstractMap.SimpleEntry<String, Double> item) {
			if (isModeFifo)
				store.addLast(item);
			else
				store.addFirst(item);
		}
	}
	
	private static DataCell generateCell(double value, String targetTypeName) {
		switch (targetTypeName) {
		case "Number (long)":
			return new LongCell((long)(value+0.4));
		case "Number (integer)":
			return new IntCell((int)(value+0.4));
		default:
			return new DoubleCell(value);
		}
	}
}
