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

package com.continental.knime.xlsformatter.cellformatter;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.knime.core.node.InvalidSettingsException;
import org.knime.core.node.NodeSettingsRO;
import org.knime.core.node.NotConfigurableException;
import org.knime.core.node.defaultnodesettings.SettingsModelString;
import org.knime.core.node.port.PortObjectSpec;
import com.continental.knime.xlsformatter.porttype.XlsFormatterState;

/**
 * SettingModelString implementation specific to CellFormatter's textRotationAngle, as we
 * experienced some undesired UTF-8 encoding/decoding differences between operating
 * systems.
 */
public class XlsFormatterCellFormatterSettingsModelString extends SettingsModelString {

	public XlsFormatterCellFormatterSettingsModelString(String configName, String defaultValue) {
		super(configName, defaultValue);
	}

	/**
   * {@inheritDoc}
   */
  @Override
  protected void loadSettingsForDialog(final NodeSettingsRO settings,
          final PortObjectSpec[] specs) throws NotConfigurableException {
  	super.loadSettingsForDialog(settings, specs);
  	updateValueWithSpecificUnicodeCorrection();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  protected void loadSettingsForModel(final NodeSettingsRO settings)
          throws InvalidSettingsException {
  	super.loadSettingsForModel(settings);
  	updateValueWithSpecificUnicodeCorrection();
  }
  
  protected Pattern regexPattern = null;
  
  /**
   * To be safe for cross-platform exchange of workflows, make sure the degree sign (�) is read in
   * correctly from the settings.xml file.
   */
  private void updateValueWithSpecificUnicodeCorrection() {
  	if (!getStringValue().equals(XlsFormatterState.FormattingFlag.UNMODIFIED.toString().toLowerCase())) {
	  	if (regexPattern == null)
	  		regexPattern = Pattern.compile("^((?:\\+|-)?(?:[0-9]{1,2}))(.*?)$");
			Matcher match = regexPattern.matcher(getStringValue());
			if (match.find())
				if (match.group(2) != null && !match.group(2).equals("°"))
					setStringValue(match.group(1) + "°");
		}
  }
}
