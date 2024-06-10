// The Data tools are a suite of ArcGIS Pro addins used to extract
// and manage biodiversity information from ArcGIS Pro and SQL Server
// based on pre-defined or user specified criteria.
//
// Copyright © 2024 Andy Foy Consulting.
//
// This file is part of DataSearches.
//
// DataSearches is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataSearches is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with DataSearches.  If not, see <http://www.gnu.org/licenses/>.

using ArcGIS.Desktop.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Xml;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

//This configuration file reader loads all of the variables to
// be used by the tool. Some are mandatory, the remainder optional.

namespace DataSearches
{
    /// <summary>
    /// This class reads the config XML file and stores the results.
    /// </summary>
    internal class DataSearchesConfig
    {
        #region Fields

        private static string _toolName;

        // Initialise component to read XML
        private readonly XmlElement _xmlDataSearches;

        #endregion Fields

        #region Constructor

        /// <summary>
        /// Load the XML profile and read the variables.
        /// </summary>
        /// <param name="xmlFile"></param>
        public DataSearchesConfig(string xmlFile, string toolName, bool msgErrors)
        {
            _toolName = toolName;

            // The user has specified the xmlFile and we've checked it exists.
            _xmlFound = true;
            _xmlLoaded = true;

            // Load the XML file into memory.
            XmlDocument xmlConfig = new();
            try
            {
                xmlConfig.Load(xmlFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading XML file. " + ex.Message, _toolName, MessageBoxButton.OK, MessageBoxImage.Error);
                _xmlLoaded = false;
                return;
            }

            // Get the InitialConfig node (the first node).
            XmlNode currNode = xmlConfig.DocumentElement.FirstChild;
            _xmlDataSearches = (XmlElement)currNode;

            if (_xmlDataSearches == null)
            {
                MessageBox.Show("Error loading XML file.", _toolName, MessageBoxButton.OK, MessageBoxImage.Error);
                _xmlLoaded = false;
                return;
            }

            // Get the mandatory variables.
            try
            {
                if (!GetMandatoryVariables())
                    return;
            }
            catch (Exception ex)
            {
                // Only report message if user was prompted for the XML
                // file (i.e. the user interface has already loaded).
                if (msgErrors)
                    MessageBox.Show("Error loading XML file. " + ex.Message, _toolName, MessageBoxButton.OK, MessageBoxImage.Error);
                _xmlLoaded = false;
                return;
            }

            // Get the map variables.
            try
            {
                if (!GetMapVariables())
                    return;
            }
            catch (Exception ex)
            {
                // Only report message if user was prompted for the XML
                // file (i.e. the user interface has already loaded).
                if (msgErrors)
                    MessageBox.Show("Error loading XML file. " + ex.Message, _toolName, MessageBoxButton.OK, MessageBoxImage.Error);
                _xmlLoaded = false;
                return;
            }
        }

        /// <summary>
        /// Get the mandatory variables from the XML file.
        /// </summary>
        /// <returns></returns>
        public bool GetMandatoryVariables()
        {
            string strRawText;

            // The existing file location where log files will be saved with output messages.
            try
            {
                _database = _xmlDataSearches["Database"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'Database' in the XML profile.");
            }

            // The field name of the search reference unique value.
            try
            {
                _refColumn = _xmlDataSearches["RefColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'RefColumn' in the XML profile.");
            }

            // The field name of the search reference site name.
            try
            {
                _siteColumn = _xmlDataSearches["SiteColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SiteColumn' in the XML profile.");
            }

            // The field name of the search reference organisation.
            try
            {
                _orgColumn = _xmlDataSearches["OrgColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'OrgColumn' in the XML profile.");
            }

            // The field name of the search reference radius.
            try
            {
                _radiusColumn = _xmlDataSearches["RadiusColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'RadiusColumn' in the XML profile.");
            }

            // Is a site name required?
            try
            {
                _requireSiteName = false;
                strRawText = _xmlDataSearches["RequireSiteName"].InnerText;
                if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _requireSiteName = true;
            }
            catch
            {
                throw new("Could not locate item 'RequireSiteName' in the XML profile.");
            }

            // The character(s) used to replace any special characters in folder names. Space is allowed.
            try
            {
                _repChar = _xmlDataSearches["RepChar"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'RepChar' in the XML profile.");
            }

            // The folder where the layer files are stored.
            try
            {
                _layerFolder = _xmlDataSearches["LayerFolder"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'LayerFolder' in the XML profile.");
            }

            // The file location where all data search folders are stored.
            try
            {
                _saveRootDir = _xmlDataSearches["SaveRootDir"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SaveRootDir' in the XML profile.");
            }

            // The folder where the report will be saved.
            try
            {
                _saveFolder = _xmlDataSearches["SaveFolder"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SaveFolder' in the XML profile.");
            }

            // The sub-folder where all data search extracts will be written to.
            try
            {
                _gisFolder = _xmlDataSearches["GISFolder"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'GISFolder' in the XML profile.");
            }

            // The log file name created by the tool to output messages.
            try
            {
                _logFileName = _xmlDataSearches["LogFileName"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'LogFileName' in the XML profile.");
            }

            // By default, should an existing log file be cleared?
            try
            {
                _defaultClearLogFile = false;
                strRawText = _xmlDataSearches["DefaultClearLogFile"].InnerText;
                if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _defaultClearLogFile = true;
            }
            catch
            {
                // This is an optional node
                _defaultClearLogFile = false;
            }

            // By default, should the log file be opened after running?
            try
            {
                _defaultOpenLogFile = false;
                strRawText = _xmlDataSearches["DefaultOpenLogFile"].InnerText;
                if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _defaultOpenLogFile = true;
            }
            catch
            {
                // This is an optional node
                _defaultOpenLogFile = false;
            }

            // The default size to use for the buffer.
            try
            {
                strRawText = _xmlDataSearches["DefaultBufferSize"].InnerText;
                bool blResult = Double.TryParse(strRawText, out double i);
                if (blResult)
                    _defaultBufferSize = (int)i;
                else
                {
                    throw new("The entry for 'DefaultBufferSize' in the XML document is not a number.");
                }
            }
            catch
            {
                // This is an optional node
                _defaultBufferSize = 0;
            }

            // The options for the buffer units. It is not recommended that these are changed.
            try
            {
                strRawText = _xmlDataSearches["BufferUnitOptions"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'BufferUnitOptions' in the XML profile.");
            }
            try
            {
                char[] chrSplit1Chars = ['$'];
                string[] liRawList = strRawText.Split(chrSplit1Chars);

                char[] chrSplit2Chars = [';'];
                foreach (string strEntry in liRawList)
                {
                    string[] strSplitEntry = strEntry.Split(chrSplit2Chars);
                    BufferUnitOptionsDisplay.Add(strSplitEntry[0]);
                    BufferUnitOptionsProcess.Add(strSplitEntry[1]);
                    BufferUnitOptionsShort.Add(strSplitEntry[2]);
                }
            }
            catch
            {
                throw new("Error parsing 'BufferUnitOptions' string. Check for correct string formatting and placement of delimiters.");
            }

            // The default option (position in the list) to use for the buffer units.
            try
            {
                strRawText = _xmlDataSearches["DefaultBufferUnit"].InnerText;
                bool blResult = Double.TryParse(strRawText, out double i);
                if (blResult)
                    _defaultBufferUnit = (int)i;
                else
                {
                    throw new("The entry for 'DefaultBufferUnit' in the XML document is not a number.");
                }
            }
            catch
            {
                // This is an optional node
                _defaultBufferUnit = -1;
            }

            // Whether the search table should be updated.
            try
            {
                _updateTable = false;
                strRawText = _xmlDataSearches["UpdateTable"].InnerText;
                if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _updateTable = true;
            }
            catch
            {
                throw new("Could not locate item 'UpdateTable' in the XML profile.");
            }

            // Are we keeping the buffer GIS file? Yes/No.
            try
            {
                _keepBufferArea = false;
                strRawText = _xmlDataSearches["KeepBufferArea"].InnerText;
                if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _keepBufferArea = true;
            }
            catch
            {
                throw new("Could not locate item 'KeepBufferArea' in the XML profile.");
            }

            // The output name for the buffer GIS file. The size of the buffer will be added automatically.
            try
            {
                _bufferSaveName = _xmlDataSearches["BufferSaveName"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'BufferSaveName' in the XML profile.");
            }

            // The name of the buffer symbology layer file.
            try
            {
                _bufferLayerName = _xmlDataSearches["BufferLayerName"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'BufferLayerName' in the XML profile.");
            }

            // The base name of the layer to use as the search area.
            try
            {
                _searchLayer = _xmlDataSearches["SearchLayer"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SearchLayer' in the XML profile.");
            }

            // The extension names for point, polygon and line search area layers.
            try
            {
                strRawText = _xmlDataSearches["SearchLayerExtensions"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SearchLayerExtensions' in the XML profile.");
            }
            try
            {
                char[] chrSplit1Chars = [';'];
                string[] liRawList = strRawText.Split(chrSplit1Chars);
                foreach (string strEntry in liRawList)
                {
                    SearchLayerExtensions.Add(strEntry);
                }
            }
            catch
            {
                throw new("Error parsing 'SearchLayerExtensions' string. Check for correct string formatting and placement of delimiters");
            }

            // The column name in the search area layer used to store the search reference.
            try
            {
                _searchColumn = _xmlDataSearches["SearchColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SearchColumn' in the XML profile.");
            }

            // Are we keeping the search feature as a layer? Yes/No.
            try
            {
                _keepSearchFeature = false;
                strRawText = _xmlDataSearches["KeepSearchFeature"].InnerText;
                if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _keepSearchFeature = true;
            }
            catch
            {
                throw new("Could not locate item 'KeepSearchFeature' in the XML profile.");
            }

            // The name of the search feature output layer.
            try
            {
                _searchFeatureName = _xmlDataSearches["SearchFeatureName"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SearchFeatureName' in the XML profile.");
            }

            // The base name of the search layer symbology file (without the .lyr).
            try
            {
                _searchSymbologyBase = _xmlDataSearches["SearchSymbologyBase"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'SearchSymbologyBase' in the XML profile.");
            }

            // The buffer aggregate column values. Delimited with semicolons.
            try
            {
                _aggregateColumns = _xmlDataSearches["AggregateColumns"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'AggregateColumns' in the XML profile.");
            }

            // The options for showing the selected tables.
            try
            {
                strRawText = _xmlDataSearches["AddSelectedLayersOptions"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'AddSelectedLayersOptions' in the XML profile.");
            }
            try
            {
                char[] chrSplitChars = [';'];
                _addSelectedLayersOptions = [.. strRawText.Split(chrSplitChars)];
            }
            catch
            {
                throw new("Error parsing 'AddSelectedLayersOptions' string. Check for correct string formatting and placement of delimiters");
            }

            // The default option (position in the list) for whether selected map layers should be added to the map window.
            try
            {
                strRawText = _xmlDataSearches["DefaultAddSelectedLayers"].InnerText;
                bool blResult = Double.TryParse(strRawText, out double i);
                if (blResult)
                    _defaultAddSelectedLayers = (int)i;
                else
                {
                    throw new("The entry for 'DefaultAddSelectedLayers' in the XML document is not a number.");
                }
            }
            catch
            {
                // This is an optional node
                _defaultAddSelectedLayers = -1;
            }

            // The name of the group layer that will be created in the ArcGIS table of contents.
            try
            {
                _groupLayerName = _xmlDataSearches["GroupLayerName"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'GroupLayerName' in the XML profile.");
            }

            // The options for overwritting the map labels.
            try
            {
                strRawText = _xmlDataSearches["OverwriteLabelOptions"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'OverwriteLabelOptions' in the XML profile.");
            }
            try
            {
                char[] chrSplitChars = [';'];
                _overwriteLabelOptions = [.. strRawText.Split(chrSplitChars)];
            }
            catch
            {
                throw new("Error parsing 'OverwriteLabelOptions' string. Check for correct string formatting and placement of delimiters");
            }

            // Whether any map label columns should be overwritten (default setting).
            try
            {
                strRawText = _xmlDataSearches["DefaultOverwriteLabels"].InnerText;
                bool blResult = Double.TryParse(strRawText, out double i);
                if (blResult)
                    _defaultOverwriteLabels = (int)i;
                else
                {
                    throw new("The entry for 'DefaultOverwriteLabels' in the XML document is not a number.");
                }
            }
            catch
            {
                // This is an optional node
                _defaultOverwriteLabels = -1;
            }

            // The units any area measurements will be done in. Choose from Ha, Km2, m2. Default is Ha..
            try
            {
                _areaMeasurementUnit = _xmlDataSearches["AreaMeasurementUnit"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'AreaMeasurementUnit' in the XML profile.");
            }

            // Options for filling out the Combined Sites table dropdown (do not change).
            try
            {
                strRawText = _xmlDataSearches["CombinedSitesTableOptions"].InnerText;
            }
            catch
            {
                throw new("Could not locate item 'CombinedSitesTableOptions' in the XML profile.");
            }
            try
            {
                char[] chrSplitChars = [';'];
                _combinedSitesTableOptions = [.. strRawText.Split(chrSplitChars)];
            }
            catch
            {
                throw new("Error parsing 'CombinedSitesTableOptions' string. Check for correct string formatting and placement of delimiters");
            }

            // Whether a combined sites table should be created by default.
            try
            {
                strRawText = _xmlDataSearches["DefaultCombinedSitesTable"].InnerText;
                bool blResult = Double.TryParse(strRawText, out double i);
                if (blResult)
                    _defaultCombinedSitesTable = (int)i;
                else
                {
                    throw new("The entry for 'DefaultCombinedSitesTable' in the XML document is not a number.");
                }
            }
            catch
            {
                // This is an optional node
                _defaultCombinedSitesTable = -1;
            }

            // The name of the combined sites table.
            try
            {
                _combinedSitesTableName = _xmlDataSearches["CombinedSitesTable"]["Name"].InnerText;
            }
            catch
            {
                throw new("Could not locate the item 'Name' for entry 'CombinedSitesTable' in the XML profile.");
            }

            // The columns of the combined sites table.
            try
            {
                _combinedSitesTableColumns = _xmlDataSearches["CombinedSitesTable"]["Columns"].InnerText;
            }
            catch
            {
                throw new("Could not locate the item 'Columns' for entry 'CombinedSitesTable' in the XML profile.");
            }

            // The format of the combined sites table.
            try
            {
                _combinedSitesTableFormat = _xmlDataSearches["CombinedSitesTable"]["Format"].InnerText;
            }
            catch
            {
                throw new("Could not locate the item 'Format' for entry 'CombinedSitesTable' in the XML profile.");
            }

            // All mandatory variables were loaded successfully.
            return true;
        }

        /// <summary>
        /// Get the map variables from the XML file.
        /// </summary>
        public bool GetMapVariables()
        {
            string strRawText;

            // The the map layer collection.
            XmlElement MapLayerCollection;
            try
            {
                MapLayerCollection = _xmlDataSearches["MapLayers"];
            }
            catch
            {
                throw new("Could not locate the item 'MapLayers' in the XML profile");
            }

            // Now cycle through all of the maps.
            if (MapLayerCollection != null)
            {
                foreach (XmlNode aNode in MapLayerCollection)
                {
                    // Only process if not a comment
                    if (aNode.NodeType != XmlNodeType.Comment)
                    {
                        string strName = aNode.Name;
                        strName = strName.Replace("_", " "); // Replace any underscores with spaces for better display.
                        MapLayers.Add(strName);
                        try
                        {
                            MapNames.Add(aNode["LayerName"].InnerText);
                        }
                        catch
                        {
                            throw new("Could not locate the item 'LayerName' for map layer " + strName + " in the XML file");
                        }

                        try
                        {
                            MapGISOutNames.Add(aNode["GISOutputName"].InnerText);
                        }
                        catch
                        {
                            throw new("Could not locate the item 'GISOutputName' for map layer " + strName + " in the XML file");
                        }

                        try
                        {
                            MapTableOutNames.Add(aNode["TableOutputName"].InnerText);
                        }
                        catch
                        {
                            throw new("Could not locate the item 'TableOutputName' for map layer " + strName + " in the XML file");
                        }

                        try
                        {
                            MapColumns.Add(aNode["Columns"].InnerText);
                        }
                        catch
                        {
                            throw new("Could not locate the item 'Columns' for map layer " + strName + " in the XML file");
                        }

                        try
                        {
                            string strGroupColumns = aNode["GroupColumns"].InnerText;
                            // Replace the commas and any spaces.
                            strGroupColumns = StringFunctions.GetGroupColumnsFormatted(strGroupColumns);
                            MapGroupColumns.Add(strGroupColumns);
                        }
                        catch
                        {
                            throw new("Could not locate the item 'GroupColumns' for map layer " + strName + " in the XML file");
                        }

                        try
                        {
                            string strStatsColumns = aNode["StatisticsColumns"].InnerText;
                            // Format the string
                            if (strStatsColumns != null)
                                strStatsColumns = StringFunctions.GetStatsColumnsFormatted(strStatsColumns);
                            MapStatisticsColumns.Add(strStatsColumns);

                        }
                        catch
                        {
                            throw new("Could not locate the item 'StatisticsColumns' for map layer " + strName + " in the XML file");
                        }

                        try
                        {
                            MapOrderColumns.Add(aNode["OrderColumns"].InnerText); // May need to deal with.
                        }
                        catch
                        {
                            throw new("Could not locate the item 'OrderColumns' for map layer " + strName + " in the XML file");
                        }

                        try
                        {
                            MapCriteria.Add(aNode["Criteria"].InnerText);
                        }
                        catch
                        {
                            throw new("Could not locate the item 'Criteria' for map layer " + strName + " in the XML file");
                        }

                        try
                        {
                            bool blIncludeArea = false;
                            strRawText = aNode["IncludeArea"].InnerText;
                            if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                blIncludeArea = true;

                            MapIncludeAreas.Add(blIncludeArea);
                        }
                        catch
                        {
                            // This is an optional node
                            MapIncludeAreas.Add(false);
                        }


                        try
                        {
                            bool blIncludDistance = false;
                            strRawText = aNode["IncludeDistance"].InnerText;
                            if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                blIncludDistance = true;

                            MapIncludeDistances.Add(blIncludDistance);
                        }
                        catch
                        {
                            // This is an optional node
                            MapIncludeDistances.Add(false);
                        }

                        try
                        {
                            bool blIncludeRadius = false;
                            strRawText = aNode["IncludeRadius"].InnerText;
                            if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                blIncludeRadius = true;

                            MapIncludeRadii.Add(blIncludeRadius);
                        }
                        catch
                        {
                            // This is an optional node
                            MapIncludeRadii.Add(false);
                        }

                        try
                        {
                            MapKeyColumns.Add(aNode["KeyColumn"].InnerText);
                        }
                        catch
                        {
                            throw new("Could not locate the item 'KeyColumn' for map layer " + strName + " in the XML file");
                        }

                        try
                        {
                            MapFormats.Add(aNode["Format"].InnerText);
                        }
                        catch
                        {
                            throw new("Could not locate the item 'Format' for map layer " + strName + " in the XML file");
                        }

                        try
                        {
                            bool blKeepLayer = false;
                            strRawText = aNode["KeepLayer"].InnerText;
                            if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                blKeepLayer = true;

                            MapKeepLayers.Add(blKeepLayer);
                        }
                        catch
                        {
                            // This is an optional node
                            MapKeepLayers.Add(false);
                        }

                        try
                        {
                            string strMapOutput = aNode["OutputType"].InnerText;
                            string strOutputType;
                            switch (strMapOutput.ToLower(System.Globalization.CultureInfo.CurrentCulture))
                            {
                                case "copy":
                                    strOutputType = "COPY";
                                    break;

                                case "clip":
                                    strOutputType = "CLIP";
                                    break;

                                case "overlay":
                                    strOutputType = "OVERLAY";
                                    break;

                                case "intersect":
                                    strOutputType = "INTERSECT";
                                    break;

                                default:
                                    strOutputType = "COPY";
                                    break;
                            }

                            MapOutputTypes.Add(strOutputType);
                        }
                        catch
                        {
                            throw new("Could not locate the item 'OutputType' for map layer " + strName + " in the XML file");
                        }

                        try
                        {
                            bool blLoadWarning = false;
                            strRawText = aNode["LoadWarning"].InnerText;
                            if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                blLoadWarning = true;

                            MapLoadWarnings.Add(blLoadWarning);
                        }
                        catch
                        {
                            // This is an optional node
                            MapLoadWarnings.Add(false);
                        }

                        try
                        {
                            bool blPreselectLayer = false;
                            strRawText = aNode["PreselectLayer"].InnerText;
                            if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                blPreselectLayer = true;

                            MapPreselectLayers.Add(blPreselectLayer);
                        }
                        catch
                        {
                            // This is an optional node
                            MapPreselectLayers.Add(false);
                        }

                        try
                        {
                            bool blDisplayLabels = false;
                            strRawText = aNode["DisplayLabels"].InnerText;
                            if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                blDisplayLabels = true;

                            MapDisplayLabels.Add(blDisplayLabels);
                        }
                        catch
                        {
                            // This is an optional node
                            MapDisplayLabels.Add(false);
                        }

                        try
                        {
                            MapLayerFiles.Add(aNode["LayerFileName"].InnerText);
                        }
                        catch
                        {
                            // This is an optional node
                            MapLayerFiles.Add("");
                        }

                        try
                        {
                            bool blOverwriteLabels = false;
                            strRawText = aNode["OverwriteLabels"].InnerText;
                            if (strRawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                blOverwriteLabels = true;

                            MapOverwriteLabels.Add(blOverwriteLabels);
                        }
                        catch
                        {
                            // This is an optional node
                            MapOverwriteLabels.Add(false);
                        }

                        try
                        {
                            MapLabelColumns.Add(aNode["LabelColumn"].InnerText);
                        }
                        catch
                        {
                            // This is an optional node
                            MapLabelColumns.Add("");
                        }

                        try
                        {
                            MapLabelClauses.Add(aNode["LabelClause"].InnerText);
                        }
                        catch
                        {
                            // This is an optional node
                            MapLabelClauses.Add("");
                        }

                        try
                        {
                            MapMacroNames.Add(aNode["MacroName"].InnerText);
                        }
                        catch
                        {
                            // This is an optional node
                            MapMacroNames.Add("");
                        }

                        bool blCombinedSites = false;
                        try
                        {
                            string strSitesColumns = aNode["CombinedSitesColumns"].InnerText;
                            if (strSitesColumns != "")
                            {
                                blCombinedSites = true;
                                MapCombinedSitesColumns.Add(strSitesColumns);
                            }
                            else
                            {
                                MapCombinedSitesColumns.Add("");
                                MapCombinedSitesGroupColumns.Add("");
                                MapCombinedSitesStatsColumns.Add("");
                                MapCombinedSitesOrderColumns.Add("");
                            }
                        }
                        catch
                        {
                            // This is an optional node
                            MapCombinedSitesColumns.Add("");
                            MapCombinedSitesGroupColumns.Add("");
                            MapCombinedSitesStatsColumns.Add("");
                            MapCombinedSitesOrderColumns.Add("");
                        }

                        // If there are any combined sites columns get the other settings
                        if (blCombinedSites)
                        {
                            try
                            {
                                string strGroupColumns = aNode["CombinedSitesGroupColumns"].InnerText;
                                // Replace delimiters
                                strGroupColumns = StringFunctions.GetGroupColumnsFormatted(strGroupColumns);
                                MapCombinedSitesGroupColumns.Add(strGroupColumns);
                            }
                            catch
                            {
                                throw new("Could not locate the item 'CombinedSitesGroupColumns' for map layer " + strName + " in the XML file");
                            }

                            try
                            {
                                string strStatsColumns = aNode["CombinedSitesStatisticsColumns"].InnerText;
                                // Format the string
                                strStatsColumns = StringFunctions.GetStatsColumnsFormatted(strStatsColumns);
                                MapCombinedSitesStatsColumns.Add(strStatsColumns);
                            }
                            catch
                            {
                                throw new("Could not locate the item 'CombinedSitesStatisticsColumns' for map layer " + strName + " in the XML file");
                            }

                            try
                            {
                                MapCombinedSitesOrderColumns.Add(aNode["CombinedSitesOrderByColumns"].InnerText); // May need to deal.
                            }
                            catch
                            {
                                throw new("Could not locate the item 'CombinedSitesOrderByColumns' for map layer " + strName + " in the XML file");
                            }
                        }
                    }
                }
            }

            // All mandatory variables were loaded successfully.
            return true;
        }

        #endregion Constructor

        #region Members

        private readonly bool _xmlFound;

        /// <summary>
        /// Has the XML file been found.
        /// </summary>
        public bool XMLFound
        {
            get
            {
                return _xmlFound;
            }
        }

        private readonly bool _xmlLoaded;

        /// <summary>
        ///  Has the XML file been loaded.
        /// </summary>
        public bool XMLLoaded
        {
            get
            {
                return _xmlLoaded;
            }
        }

        #endregion Members

        #region Variables

        private string _database;

        public string Database
        {
            get { return _database; }
        }

        private string _refColumn;

        public string RefColumn
        {
            get { return _refColumn; }
        }

        private string _siteColumn;

        public string SiteColumn
        {
            get { return _siteColumn; }
        }

        private string _orgColumn;

        public string OrgColumn
        {
            get { return _orgColumn; }
        }

        private string _radiusColumn;

        public string RadiusColumn
        {
            get { return _radiusColumn; }
        }

        private bool _requireSiteName;

        public bool RequireSiteName
        {
            get { return _requireSiteName; }
        }

        private string _repChar;

        public string RepChar
        {
            get { return _repChar; }
        }

        private string _layerFolder;

        public string LayerFolder
        {
            get { return _layerFolder; }
        }

        private string _saveRootDir;

        public string SaveRootDir
        {
            get { return _saveRootDir; }
        }

        private string _saveFolder;

        public string SaveFolder
        {
            get { return _saveFolder; }
        }

        private string _gisFolder;

        public string GisFolder
        {
            get { return _gisFolder; }
        }

        private string _logFileName;

        public string LogFileName
        {
            get { return _logFileName; }
        }

        private bool _defaultClearLogFile;

        public bool DefaultClearLogFile
        {
            get { return _defaultClearLogFile; }
        }

        private bool _defaultOpenLogFile;

        public bool DefaultOpenLogFile
        {
            get { return _defaultOpenLogFile; }
        }

        private int _defaultBufferSize;

        public int DefaultBufferSize
        {
            get { return _defaultBufferSize; }
        }

        private List<string> _bufferUnitOptionsDisplay = [];

        public List<string> BufferUnitOptionsDisplay
        {
            get { return _bufferUnitOptionsDisplay; }
        }

        private List<string> _bufferUnitOptionsProcess = [];

        public List<string> BufferUnitOptionsProcess
        {
            get { return _bufferUnitOptionsProcess; }
        }

        private List<string> _bufferUnitOptionsShort = [];

        public List<string> BufferUnitOptionsShort
        {
            get { return _bufferUnitOptionsShort; }
        }

        private int _defaultBufferUnit;

        public int DefaultBufferUnit
        {
            get { return _defaultBufferUnit; }
        }

        private bool _updateTable;

        public bool UpdateTable
        {
            get { return _updateTable; }
        }

        private bool _keepBufferArea;

        public bool KeepBufferArea
        {
            get { return _keepBufferArea; }
        }

        private string _bufferSaveName;

        public string BufferSaveName
        {
            get { return _bufferSaveName; }
        }

        private string _bufferLayerName;

        public string BufferLayerName
        {
            get { return _bufferLayerName; }
        }

        private string _searchLayer;

        public string SearchLayer
        {
            get { return _searchLayer; }
        }

        private List<string> _searchLayerExtensions = [];

        public List<string> SearchLayerExtensions
        {
            get { return _searchLayerExtensions; }
        }

        private string _searchColumn;

        public string SearchColumn
        {
            get { return _searchColumn; }
        }

        private bool _keepSearchFeature;

        public bool KeepSearchFeature
        {
            get { return _keepSearchFeature; }
        }

        private string _searchFeatureName;

        public string SearchFeatureName
        {
            get { return _searchFeatureName; }
        }

        private string _searchSymbologyBase;

        public string SearchSymbologyBase
        {
            get { return _searchSymbologyBase; }
        }

        private string _aggregateColumns;

        public string AggregateColumns
        {
            get { return _aggregateColumns; }
        }

        private List<string> _addSelectedLayersOptions = [];

        public List<string> AddSelectedLayersOptions
        {
            get { return _addSelectedLayersOptions; }
        }

        private int _defaultAddSelectedLayers;

        public int DefaultAddSelectedLayers
        {
            get { return _defaultAddSelectedLayers; }
        }

        private string _groupLayerName;

        public string GroupLayerName
        {
            get { return _groupLayerName; }
        }

        private List<string> _overwriteLabelOptions = [];

        public List<string> OverwriteLabelOptions
        {
            get { return _overwriteLabelOptions; }
        }

        private int _defaultOverwriteLabels;

        public int DefaultOverwriteLabels
        {
            get { return _defaultOverwriteLabels; }
        }

        private string _areaMeasurementUnit;

        public string AreaMeasurementUnit
        {
            get { return _areaMeasurementUnit; }
        }

        private List<string> _combinedSitesTableOptions = [];

        public List<string> CombinedSitesTableOptions
        {
            get { return _combinedSitesTableOptions; }
        }

        private int _defaultCombinedSitesTable;

        public int DefaultCombinedSitesTable
        {
            get { return _defaultCombinedSitesTable; }
        }

        private string _combinedSitesTableName;

        public string CombinedSitesTableName
        {
            get { return _combinedSitesTableName; }
        }

        private string _combinedSitesTableColumns;

        public string CombinedSitesTableColumns
        {
            get { return _combinedSitesTableColumns; }
        }

        private string _combinedSitesTableFormat;

        public string CombinedSitesTableFormat
        {
            get { return _combinedSitesTableFormat; }
        }

        #endregion

        #region Map Variables

        private List<string> _mapLayers = [];

        public List<string> MapLayers
        {
            get { return _mapLayers; }
        }

        private List<string> _mapNames = [];

        public List<string> MapNames
        {
            get { return _mapNames; }
        }

        private List<string> _mapGISOutNames = [];

        public List<string> MapGISOutNames
        {
            get { return _mapGISOutNames; }
        }

        private List<string> _mapTableOutNames = [];

        public List<string> MapTableOutNames
        {
            get { return _mapTableOutNames; }
        }

        private List<string> _mapColumns = [];

        public List<string> MapColumns
        {
            get { return _mapColumns; }
        }

        private List<string> _mapGroupColumns = [];

        public List<string> MapGroupColumns
        {
            get { return _mapGroupColumns; }
        }

        private List<string> _mapStatisticsColumns = [];

        public List<string> MapStatisticsColumns
        {
            get { return _mapStatisticsColumns; }
        }

        private List<string> _mapOrderColumns = [];

        public List<string> MapOrderColumns
        {
            get { return _mapOrderColumns; }
        }

        private List<string> _mapCriteria = [];

        public List<string> MapCriteria
        {
            get { return _mapCriteria; }
        }

        private List<bool> _mapIncludeAreas = [];

        public List<bool> MapIncludeAreas
        {
            get { return _mapIncludeAreas; }
        }

        private List<bool> _mapIncludeDistances = [];

        public List<bool> MapIncludeDistances
        {
            get { return _mapIncludeDistances; }
        }

        private List<bool> _mapIncludeRadii = [];

        public List<bool> MapIncludeRadii
        {
            get { return _mapIncludeRadii; }
        }

        private List<string> _mapKeyColumns = [];

        public List<string> MapKeyColumns
        {
            get { return _mapKeyColumns; }
        }

        private List<string> _mapFormats = [];

        public List<string> MapFormats
        {
            get { return _mapFormats; }
        }

        private List<bool> _mapKeepLayers = [];

        public List<bool> MapKeepLayers
        {
            get { return _mapKeepLayers; }
        }

        private List<string> _mapOutputTypes = [];

        public List<string> MapOutputTypes
        {
            get { return _mapOutputTypes; }
        }

        private List<bool> _mapIntersectOutputs = [];

        public List<bool> MapIntersectOutputs
        {
            get { return _mapIntersectOutputs; }
        }

        private List<bool> _mapLoadWarnings = [];

        public List<bool> MapLoadWarnings
        {
            get { return _mapLoadWarnings; }
        }

        private List<bool> _mapPreselectLayers = [];

        public List<bool> MapPreselectLayers
        {
            get { return _mapPreselectLayers; }
        }

        private List<bool> _mapDisplayLabels = [];

        public List<bool> MapDisplayLabels
        {
            get { return _mapDisplayLabels; }
        }

        private List<string> _mapLayerFiles = [];

        public List<string> MapLayerFiles
        {
            get { return _mapLayerFiles; }
        }

        private List<bool> _mapOverwriteLabels = [];

        public List<bool> MapOverwriteLabels
        {
            get { return _mapOverwriteLabels; }
        }

        private List<string> _mapLabelColumns = [];

        public List<string> MapLabelColumns
        {
            get { return _mapLabelColumns; }
        }

        private List<string> _mapLabelClauses = [];

        public List<string> MapLabelClauses
        {
            get { return _mapLabelClauses; }
        }

        private List<string> _mapMacroNames = [];

        public List<string> MapMacroNames
        {
            get { return _mapMacroNames; }
        }

        private List<string> _mapCombinedSitesColumns = [];

        public List<string> MapCombinedSitesColumns
        {
            get { return _mapCombinedSitesColumns; }
        }

        private List<string> _mapCombinedSitesGroupColumns = [];

        public List<string> MapCombinedSitesGroupColumns
        {
            get { return _mapCombinedSitesGroupColumns; }
        }

        private List<string> _mapCombinedSitesStatsColumns = [];

        public List<string> MapCombinedSitesStatsColumns
        {
            get { return _mapCombinedSitesStatsColumns; }
        }

        private List<string> _mapCombinedSitesOrderColumns = [];

        public List<string> MapCombinedSitesOrderColumns
        {
            get { return _mapCombinedSitesOrderColumns; }
        }

        #endregion Variables
    }
}