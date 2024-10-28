// The DataTools are a suite of ArcGIS Pro addins used to extract
// and manage biodiversity information from ArcGIS Pro and SQL Server
// based on pre-defined or user specified criteria.
//
// Copyright © 2024 Andy Foy Consulting.
//
// This file is part of DataTools suite of programs..
//
// DataTools are free software: you can redistribute it and/or modify
// them under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataTools are distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with with program.  If not, see <http://www.gnu.org/licenses/>.

using DataSearches.UI;
using DataTools;
using System;
using System.Collections.Generic;
using System.Windows;
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
        }

        #endregion Constructor

        #region Get Mandatory Variables

        /// <summary>
        /// Get the mandatory variables from the XML file.
        /// </summary>
        /// <returns></returns>
        public bool GetMandatoryVariables()
        {
            string rawText;

            // The access database where all the data search details are stored.
            try
            {
                _databasePath = _xmlDataSearches["DatabasePath"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'DatabasePath' in the XML profile.");
            }

            // Only get the other database details if there is a path.
            if (!string.IsNullOrEmpty(_databasePath))
            {
                // The name of the table where the enquiries are stored in the database table.
                try
                {
                    _databaseTable = _xmlDataSearches["DatabaseTable"].InnerText;
                }
                catch
                {
                    throw new("Could not locate 'DatabaseTable' in the XML profile.");
                }

                // The column name of the search reference unique value in the database table.
                try
                {
                    _databaseRefColumn = _xmlDataSearches["DatabaseRefColumn"].InnerText;
                }
                catch
                {
                    throw new("Could not locate 'DatabaseRefColumn' in the XML profile.");
                }

                // The column name of the site name in the database table.
                try
                {
                    _databaseSiteColumn = _xmlDataSearches["DatabaseSiteColumn"].InnerText;
                }
                catch
                {
                    throw new("Could not locate 'DatabaseSiteColumn' in the XML profile.");
                }

                // The column name of the organisation in the database table.
                try
                {
                    _databaseOrgColumn = _xmlDataSearches["DatabaseOrgColumn"].InnerText;
                }
                catch
                {
                    throw new("Could not locate 'DatabaseOrgColumn' in the XML profile.");
                }
            }
            else
            {
                _databaseTable = null;
                _databaseRefColumn = null;
                _databaseSiteColumn = null;
                _databaseOrgColumn = null;
            }

            // Is a site name required?
            try
            {
                _requireSiteName = false;
                rawText = _xmlDataSearches["RequireSiteName"].InnerText;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _requireSiteName = true;
            }
            catch
            {
                throw new("Could not locate 'RequireSiteName' in the XML profile.");
            }

            // Is the organisation required?
            try
            {
                _requireOrganisation = false;
                rawText = _xmlDataSearches["RequireOrganisation"].InnerText;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _requireOrganisation = true;
            }
            catch
            {
                throw new("Could not locate 'RequireOrganisation' in the XML profile.");
            }

            // The character(s) used to replace any special characters in folder names. Space is allowed.
            try
            {
                _repChar = _xmlDataSearches["RepChar"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'RepChar' in the XML profile.");
            }

            // The folder where the layer files are stored.
            try
            {
                _layerFolder = _xmlDataSearches["LayerFolder"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'LayerFolder' in the XML profile.");
            }

            // The file location where all data search folders are stored.
            try
            {
                _saveRootDir = _xmlDataSearches["SaveRootDir"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'SaveRootDir' in the XML profile.");
            }

            // The folder where the report will be saved.
            try
            {
                _saveFolder = _xmlDataSearches["SaveFolder"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'SaveFolder' in the XML profile.");
            }

            // The sub-folder where all data search extracts will be written to.
            try
            {
                _gisFolder = _xmlDataSearches["GISFolder"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'GISFolder' in the XML profile.");
            }

            // The log file name created by the tool to output messages.
            try
            {
                _logFileName = _xmlDataSearches["LogFileName"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'LogFileName' in the XML profile.");
            }

            // Whether the map processing should be paused during processing?
            try
            {
                _pauseMap = false;
                rawText = _xmlDataSearches["PauseMap"].InnerText;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _pauseMap = true;
            }
            catch
            {
                // This is an optional node
                _pauseMap = false;
            }

            // By default, should an existing log file be cleared?
            try
            {
                _defaultClearLogFile = false;
                rawText = _xmlDataSearches["DefaultClearLogFile"].InnerText;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
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
                rawText = _xmlDataSearches["DefaultOpenLogFile"].InnerText;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
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
                rawText = _xmlDataSearches["DefaultBufferSize"].InnerText;
                bool blResult = Double.TryParse(rawText, out double i);
                if (blResult)
                    _defaultBufferSize = (int)i;
                else
                {
                    throw new("The entry for 'DefaultBufferSize' in the XML profile is not a number.");
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
                rawText = _xmlDataSearches["BufferUnitOptions"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'BufferUnitOptions' in the XML profile.");
            }
            try
            {
                char[] chrSplit1Chars = ['$'];
                string[] liRawList = rawText.Split(chrSplit1Chars);

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
                throw new("Error parsing 'BufferUnitOptions' string. Check for correct format.");
            }

            // The default option (position in the list) to use for the buffer units.
            try
            {
                rawText = _xmlDataSearches["DefaultBufferUnit"].InnerText;
                bool blResult = Double.TryParse(rawText, out double i);
                if (blResult)
                    _defaultBufferUnit = (int)i;
                else
                {
                    throw new("The entry for 'DefaultBufferUnit' in the XML profile is not a number.");
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
                rawText = _xmlDataSearches["UpdateTable"].InnerText;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _updateTable = true;
            }
            catch
            {
                throw new("Could not locate 'UpdateTable' in the XML profile.");
            }

            // Are we keeping the buffer GIS file? Yes/No.
            try
            {
                _keepBufferArea = false;
                rawText = _xmlDataSearches["KeepBufferArea"].InnerText;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _keepBufferArea = true;
            }
            catch
            {
                throw new("Could not locate 'KeepBufferArea' in the XML profile.");
            }

            // The prefix output name for the buffer GIS file. The size of the buffer will be added automatically.
            try
            {
                _bufferPrefix = _xmlDataSearches["BufferPrefix"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'BufferPrefix' in the XML profile.");
            }

            // The name of the buffer symbology layer file.
            try
            {
                _bufferLayerFile = _xmlDataSearches["BufferLayerFile"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'BufferLayerFile' in the XML profile.");
            }

            // The base name of the layer to use as the search area.
            try
            {
                _searchLayer = _xmlDataSearches["SearchLayer"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'SearchLayer' in the XML profile.");
            }

            // The extension names for point, polygon and line search area layers.
            try
            {
                rawText = _xmlDataSearches["SearchLayerExtensions"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'SearchLayerExtensions' in the XML profile.");
            }
            try
            {
                char[] chrSplit1Chars = [';'];
                string[] liRawList = rawText.Split(chrSplit1Chars);
                foreach (string rawEntry in liRawList)
                {
                    SearchLayerExtensions.Add(rawEntry);
                }
            }
            catch
            {
                throw new("Error parsing 'SearchLayerExtensions' string. Check for correct format.");
            }

            // The column name in the search area layer used to store the search reference.
            try
            {
                _searchColumn = _xmlDataSearches["SearchColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'SearchColumn' in the XML profile.");
            }

            // The column name in the search area layer used to store the site name.
            try
            {
                _siteColumn = _xmlDataSearches["SiteColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'SiteColumn' in the XML profile.");
            }

            // The column name in the search area layer used to store the organisation.
            try
            {
                _orgColumn = _xmlDataSearches["OrgColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'OrgColumn' in the XML profile.");
            }

            // The column name in the search area layer used to store the radius.
            try
            {
                _radiusColumn = _xmlDataSearches["RadiusColumn"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'RadiusColumn' in the XML profile.");
            }

            // Are we keeping the search feature as a layer? Yes/No.
            try
            {
                _keepSearchFeature = false;
                rawText = _xmlDataSearches["KeepSearchFeature"].InnerText;
                if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _keepSearchFeature = true;
            }
            catch
            {
                throw new("Could not locate 'KeepSearchFeature' in the XML profile.");
            }

            // The name of the search feature output layer.
            try
            {
                _searchOutputName = _xmlDataSearches["SearchOutputName"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'SearchOutputName' in the XML profile.");
            }

            // The base name of the search layer symbology file (without the .lyr).
            try
            {
                _searchSymbologyBase = _xmlDataSearches["SearchSymbologyBase"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'SearchSymbologyBase' in the XML profile.");
            }

            // The buffer aggregate column values. Delimited with semicolons.
            try
            {
                _aggregateColumns = _xmlDataSearches["AggregateColumns"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'AggregateColumns' in the XML profile.");
            }

            // The options for showing the selected tables.
            try
            {
                rawText = _xmlDataSearches["AddSelectedLayersOptions"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'AddSelectedLayersOptions' in the XML profile.");
            }
            try
            {
                char[] chrSplitChars = [';'];
                _addSelectedLayersOptions = [.. rawText.Split(chrSplitChars)];
            }
            catch
            {
                throw new("Error parsing 'AddSelectedLayersOptions' string. Check for correct format.");
            }

            // The default option for whether selected map layers should be kept.
            try
            {
                _defaultKeepSelectedLayers = false;
                rawText = _xmlDataSearches["DefaultKeepSelectedLayers"].InnerText;
                if (string.IsNullOrEmpty(rawText))
                    _defaultKeepSelectedLayers = null;
                else if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                    _defaultKeepSelectedLayers = true;
            }
            catch
            {
                // This is an optional node
                _defaultKeepSelectedLayers = false;
            }

            // The default option (position in the list) for whether selected map layers should be added to the map window.
            try
            {
                rawText = _xmlDataSearches["DefaultAddSelectedLayers"].InnerText;
                bool blResult = Double.TryParse(rawText, out double i);
                if (blResult)
                    _defaultAddSelectedLayers = (int)i;
                else
                {
                    throw new("The entry for 'DefaultAddSelectedLayers' in the XML profile is not a number.");
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
                throw new("Could not locate 'GroupLayerName' in the XML profile.");
            }

            // The options for overwritting the map labels.
            try
            {
                rawText = _xmlDataSearches["OverwriteLabelOptions"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'OverwriteLabelOptions' in the XML profile.");
            }
            try
            {
                char[] chrSplitChars = [';'];
                _overwriteLabelOptions = [.. rawText.Split(chrSplitChars)];
            }
            catch
            {
                throw new("Error parsing 'OverwriteLabelOptions' string. Check for correct format.");
            }

            // Whether any map label columns should be overwritten (default setting).
            try
            {
                rawText = _xmlDataSearches["DefaultOverwriteLabels"].InnerText;
                bool blResult = Double.TryParse(rawText, out double i);
                if (blResult)
                    _defaultOverwriteLabels = (int)i;
                else
                {
                    throw new("The entry for 'DefaultOverwriteLabels' in the XML profile is not a number.");
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
                throw new("Could not locate 'AreaMeasurementUnit' in the XML profile.");
            }

            // Options for filling out the Combined Sites table dropdown (do not change).
            try
            {
                rawText = _xmlDataSearches["CombinedSitesTableOptions"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'CombinedSitesTableOptions' in the XML profile.");
            }
            try
            {
                char[] chrSplitChars = [';'];
                _combinedSitesTableOptions = [.. rawText.Split(chrSplitChars)];
            }
            catch
            {
                throw new("Error parsing 'CombinedSitesTableOptions' string. Check for correct format.");
            }

            // Whether a combined sites table should be created by default.
            try
            {
                rawText = _xmlDataSearches["DefaultCombinedSitesTable"].InnerText;
                bool blResult = Double.TryParse(rawText, out double i);
                if (blResult)
                    _defaultCombinedSitesTable = (int)i;
                else
                {
                    throw new("The entry for 'DefaultCombinedSitesTable' in the XML profile is not a number.");
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
                throw new("Could not locate 'Name' for entry 'CombinedSitesTable' in the XML profile.");
            }

            // The columns of the combined sites table.
            try
            {
                _combinedSitesTableColumns = _xmlDataSearches["CombinedSitesTable"]["Columns"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'Columns' for entry 'CombinedSitesTable' in the XML profile.");
            }

            // The format of the combined sites table.
            try
            {
                _combinedSitesTableFormat = _xmlDataSearches["CombinedSitesTable"]["Format"].InnerText;
            }
            catch
            {
                throw new("Could not locate 'Format' for entry 'CombinedSitesTable' in the XML profile.");
            }

            // All mandatory variables were loaded successfully.
            return true;
        }

        #endregion Get Mandatory Variables

        #region Get Map Variables

        /// <summary>
        /// Get the map variables from the XML file.
        /// </summary>
        public bool GetMapVariables()
        {
            string rawText;

            // The the map layer collection.
            XmlElement MapLayerCollection;
            try
            {
                MapLayerCollection = _xmlDataSearches["MapLayers"];
            }
            catch
            {
                throw new("Could not locate 'MapLayers' in the XML profile");
            }

            // Reset the map layers list.
            _mapLayers = [];

            // Now cycle through all of the maps.
            if (MapLayerCollection != null)
            {
                foreach (XmlNode node in MapLayerCollection)
                {
                    // Only process if not a comment
                    if (node.NodeType != XmlNodeType.Comment)
                    {
                        string nodeName = node.Name;

                        // Replace any underscores with spaces for better display.
                        nodeName = nodeName.Replace("_", " ");

                        // Create a new layer for this node.
                        MapLayer layer = new(nodeName);

                        try
                        {
                            string nodeGroup = nodeName.Substring(0, nodeName.IndexOf('-')).Trim();
                            string nodeLayer = nodeName.Substring(nodeName.IndexOf('-') + 1).Trim();
                            layer.NodeGroup = nodeGroup;
                            layer.NodeLayer = nodeLayer;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.NodeGroup = null;
                            layer.NodeLayer = nodeName;
                        }

                        try
                        {
                            layer.LayerName = node["LayerName"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate 'LayerName' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            layer.GISOutputName = node["GISOutputName"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate 'GISOutputName' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            layer.TableOutputName = node["TableOutputName"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate 'TableOutputName' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            layer.Columns = node["Columns"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate 'Columns' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            string groupColumns = node["GroupColumns"].InnerText;
                            // Replace the commas and any spaces.
                            groupColumns = StringFunctions.GetGroupColumnsFormatted(groupColumns);

                            layer.GroupColumns = groupColumns;
                        }
                        catch
                        {
                            throw new("Could not locate 'GroupColumns' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            string statsColumns = node["StatisticsColumns"].InnerText;
                            //// Format the string
                            //statsColumns = StringFunctions.GetStatsColumnsFormatted(statsColumns);

                            layer.StatisticsColumns = statsColumns;
                        }
                        catch
                        {
                            throw new("Could not locate 'StatisticsColumns' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            layer.OrderColumns = node["OrderColumns"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate 'OrderColumns' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            layer.Criteria = node["Criteria"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate 'Criteria' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            bool includeArea = false;
                            rawText = node["IncludeArea"].InnerText;
                            if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                includeArea = true;

                            layer.IncludeArea = includeArea;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.IncludeArea = false;
                        }

                        try
                        {
                            rawText = node["IncludeNearFields"].InnerText;
                            string includeNear = (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture)) switch
                            {
                                "centroid" => "CENTROID",
                                "boundary" => "BOUNDARY",
                                _ => "",
                            };

                            layer.IncludeNearFields = includeNear;
                        }
                        catch
                        {
                            throw new("Could not locate 'IncludeNearFields' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            bool includeRadius = false;
                            rawText = node["IncludeRadius"].InnerText;
                            if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                includeRadius = true;

                            layer.IncludeRadius = includeRadius;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.IncludeRadius = false;
                        }

                        try
                        {
                            layer.KeyColumn = node["KeyColumn"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate 'KeyColumn' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            layer.Format = node["Format"].InnerText;
                        }
                        catch
                        {
                            throw new("Could not locate 'Format' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            bool keepLayer = false;
                            rawText = node["KeepLayer"].InnerText;
                            if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                keepLayer = true;

                            layer.KeepLayer = keepLayer;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.KeepLayer = false;
                        }

                        try
                        {
                            string strMapOutput = node["OutputType"].InnerText;
                            string outputType = (strMapOutput.ToLower(System.Globalization.CultureInfo.CurrentCulture)) switch
                            {
                                "copy" => "COPY",
                                "clip" => "CLIP",
                                "overlay" => "OVERLAY",
                                "intersect" => "INTERSECT",
                                _ => "COPY",
                            };

                            layer.OutputType = outputType;
                        }
                        catch
                        {
                            throw new("Could not locate 'OutputType' for map layer '" + nodeName + "'.");
                        }

                        try
                        {
                            bool loadWarning = false;
                            rawText = node["LoadWarning"].InnerText;
                            if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                loadWarning = true;

                            layer.LoadWarning = loadWarning;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.LoadWarning = false;
                        }

                        try
                        {
                            bool preselectLayer = false;
                            rawText = node["PreselectLayer"].InnerText;
                            if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                preselectLayer = true;

                            layer.PreselectLayer = preselectLayer;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.PreselectLayer = false;
                        }

                        try
                        {
                            bool displayLabels = false;
                            rawText = node["DisplayLabels"].InnerText;
                            if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                displayLabels = true;

                            layer.DisplayLabels = displayLabels;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.DisplayLabels = false;
                        }

                        try
                        {
                            layer.LayerFileName = node["LayerFileName"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.LayerFileName = null;
                        }

                        try
                        {
                            bool overwriteLabels = false;
                            rawText = node["OverwriteLabels"].InnerText;
                            if (rawText.ToLower(System.Globalization.CultureInfo.CurrentCulture) is "yes" or "y")
                                overwriteLabels = true;

                            layer.OverwriteLabels = overwriteLabels;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.OverwriteLabels = false;
                        }

                        try
                        {
                            layer.LabelColumn = node["LabelColumn"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.LabelColumn = null;
                        }

                        try
                        {
                            layer.LabelClause = node["LabelClause"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.LabelClause = null;
                        }

                        try
                        {
                            layer.MacroName = node["MacroName"].InnerText;
                        }
                        catch
                        {
                            // This is an optional node
                            layer.MacroName = null;
                        }

                        bool combinedSites = false;
                        try
                        {
                            string sitesColumns = node["CombinedSitesColumns"].InnerText;
                            if (!string.IsNullOrEmpty(sitesColumns))
                            {
                                combinedSites = true;
                                layer.CombinedSitesColumns = node["CombinedSitesColumns"].InnerText;
                            }
                            else
                            {
                                layer.CombinedSitesColumns = null;
                                layer.CombinedSitesGroupColumns = null;
                                layer.CombinedSitesStatisticsColumns = null;
                                layer.CombinedSitesOrderByColumns = null;
                            }
                        }
                        catch
                        {
                            // This is an optional node
                            layer.CombinedSitesColumns = null;
                            layer.CombinedSitesGroupColumns = null;
                            layer.CombinedSitesStatisticsColumns = null;
                            layer.CombinedSitesOrderByColumns = null;
                        }

                        // If there are any combined sites columns get the other settings
                        if (combinedSites)
                        {
                            try
                            {
                                string groupColumns = node["CombinedSitesGroupColumns"].InnerText;
                                // Replace delimiters
                                groupColumns = StringFunctions.GetGroupColumnsFormatted(groupColumns);

                                layer.CombinedSitesGroupColumns = groupColumns;
                            }
                            catch
                            {
                                throw new("Could not locate 'CombinedSitesGroupColumns' for map layer '" + nodeName + "'.");
                            }

                            try
                            {
                                string statsColumns = node["CombinedSitesStatisticsColumns"].InnerText;
                                //// Format the string
                                //statsColumns = StringFunctions.GetStatsColumnsFormatted(statsColumns);

                                layer.CombinedSitesStatisticsColumns = statsColumns;
                            }
                            catch
                            {
                                throw new("Could not locate 'CombinedSitesStatisticsColumns' for map layer '" + nodeName + "'.");
                            }

                            try
                            {
                                layer.CombinedSitesOrderByColumns = node["CombinedSitesOrderByColumns"].InnerText;
                            }
                            catch
                            {
                                throw new("Could not locate 'CombinedSitesOrderByColumns' for map layer '" + nodeName + "'.");
                            }
                        }

                        // Add the layer to the list of map layers.
                        MapLayers.Add(layer);
                    }
                }
            }

            // All mandatory variables were loaded successfully.
            return true;
        }

        #endregion Get Map Variables

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

        private string _databasePath;

        public string DatabasePath
        {
            get { return _databasePath; }
        }

        private string _databaseTable;

        public string DatabaseTable
        {
            get { return _databaseTable; }
        }

        private string _databaseRefColumn;

        public string DatabaseRefColumn
        {
            get { return _databaseRefColumn; }
        }

        private string _databaseSiteColumn;

        public string DatabaseSiteColumn
        {
            get { return _databaseSiteColumn; }
        }

        private string _databaseOrgColumn;

        public string DatabaseOrgColumn
        {
            get { return _databaseOrgColumn; }
        }

        private bool _requireSiteName;

        public bool RequireSiteName
        {
            get { return _requireSiteName; }
        }

        private bool _requireOrganisation;

        public bool RequireOrganisation
        {
            get { return _requireOrganisation; }
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

        public string GISFolder
        {
            get { return _gisFolder; }
        }

        private string _logFileName;

        public string LogFileName
        {
            get { return _logFileName; }
        }

        private bool _pauseMap;

        public bool PauseMap
        {
            get { return _pauseMap; }
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

        private readonly List<string> _bufferUnitOptionsDisplay = [];

        public List<string> BufferUnitOptionsDisplay
        {
            get { return _bufferUnitOptionsDisplay; }
        }

        private readonly List<string> _bufferUnitOptionsProcess = [];

        public List<string> BufferUnitOptionsProcess
        {
            get { return _bufferUnitOptionsProcess; }
        }

        private readonly List<string> _bufferUnitOptionsShort = [];

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

        private string _bufferPrefix;

        public string BufferPrefix
        {
            get { return _bufferPrefix; }
        }

        private string _bufferLayerFile;

        public string BufferLayerFile
        {
            get { return _bufferLayerFile; }
        }

        private string _searchLayer;

        public string SearchLayer
        {
            get { return _searchLayer; }
        }

        private readonly List<string> _searchLayerExtensions = [];

        public List<string> SearchLayerExtensions
        {
            get { return _searchLayerExtensions; }
        }

        private string _searchColumn;

        public string SearchColumn
        {
            get { return _searchColumn; }
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

        private bool _keepSearchFeature;

        public bool KeepSearchFeature
        {
            get { return _keepSearchFeature; }
        }

        private string _searchOutputName;

        public string SearchOutputName
        {
            get { return _searchOutputName; }
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

        private bool? _defaultKeepSelectedLayers;

        public bool? DefaultKeepSelectedLayers
        {
            get { return _defaultKeepSelectedLayers; }
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

        #endregion Variables

        #region Map Variables

        private List<MapLayer> _mapLayers = [];

        public List<MapLayer> MapLayers
        {
            get { return _mapLayers; }
        }

        #endregion Map Variables
    }
}