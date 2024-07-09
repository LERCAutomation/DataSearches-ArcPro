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

using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Data.Exceptions;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using DataSearches.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.IO.Compression;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using QueryFilter = ArcGIS.Core.Data.QueryFilter;
using System.Threading;
using System.Windows.Forms;
using ArcGIS.Desktop.Layouts;
using System.Diagnostics;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ActiproSoftware.Windows.Extensions;
using System.Security.Policy;
using System.Windows.Media;
using ArcGIS.Desktop.Internal.Catalog.Wizards;
using System.Runtime.InteropServices;
using ArcGIS.Desktop.Editing;

namespace DataSearches
{
    /// <summary>
    /// This class provides ArcGIS Pro map functions.
    /// </summary>
    internal class MapFunctions
    {
        #region Fields

        private Map _activeMap;
        private MapView _activeMapView;

        #endregion Fields

        #region Constructor

        /// <summary>
        /// Set the global variables.
        /// </summary>
        public MapFunctions()
        {
            // Get the active map view (if there is one).
            _activeMapView = GetActiveMapView();

            // Set the map currently displayed in the active map view.
            if (_activeMapView != null)
                _activeMap = _activeMapView.Map;
            else
                _activeMap = null;
        }

        #endregion Constructor

        #region Properties

        /// <summary>
        /// The name of the active map.
        /// </summary>
        public string MapName
        {
            get
            {
                if (_activeMap == null)
                    return null;
                else
                    return _activeMap.Name;
            }
        }

        #endregion Properties

        #region Map

        /// <summary>
        /// Get the active map view.
        /// </summary>
        /// <returns></returns>
        internal static MapView GetActiveMapView()
        {
            // Get the active map view.
            MapView mapView = MapView.Active;
            if (mapView == null)
                return null;

            return mapView;
        }

        /// <summary>
        /// Create a new map.
        /// </summary>
        /// <param name="mapName"></param>
        /// <returns></returns>
        public async Task<string> CreateMapAsync(string mapName)
        {
            _activeMap = null;
            _activeMapView = null;

            await QueuedTask.Run(() =>
            {
                try
                {
                    // Create a new map without a base map.
                    _activeMap = MapFactory.Instance.CreateMap(mapName, basemap: Basemap.None);

                    // Create and activate new map.
                    ArcGIS.Desktop.Framework.FrameworkApplication.Panes.CreateMapPaneAsync(_activeMap, MapViewingMode.Map);
                    //var paneTask = ProApp.Panes.CreateMapPaneAsync(_activeMap, MapViewingMode.Map);
                    //paneTask.Wait();

                    // Get the active map view;
                    //_activeMapView = GetActiveMapView();

                    //Pane pane = ProApp.Panes.ActivePane;
                    //pane.Activate();
                }
                catch
                {
                    // CreateMap throws an exception if the map view wasn't created.
                    // CreateMapPaneAsync throws an exception if the map isn't created.
                }
            });

            // Get the active map view;
            _activeMapView = GetActiveMapView();

            return _activeMap.Name;
        }

        /// <summary>
        /// Add a layer to the active map.
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public async Task AddLayerToMap(string url)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    Uri uri = new(url);

                    // Check if the layer is already loaded (unlikely as the map is new)
                    Layer findLayer = _activeMap.Layers.FirstOrDefault(t => t.Name == uri.Segments.Last());

                    // If the layer is not loaded, add it.
                    if (findLayer == null)
                    {
                        Layer layer = LayerFactory.Instance.CreateLayer(uri, _activeMap);
                    }
                });
            }
            catch
            {
                // Handle Exception.
            }
        }

        /// <summary>
        /// Add a standalone layer to the active map.
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public async Task AddTableToMap(string url)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    Uri uri = new(url);

                    // Check if the layer is already loaded.
                    StandaloneTable findTable = _activeMap.StandaloneTables.FirstOrDefault(t => t.Name == uri.Segments.Last());

                    // If the layer is not loaded, add it.
                    if (findTable == null)
                    {
                        StandaloneTable table = StandaloneTableFactory.Instance.CreateStandaloneTable(uri, _activeMap);
                    }
                });
            }
            catch
            {
                // Handle Exception.
            }
        }

        #endregion Map

        #region Layers

        /// <summary>
        /// Find a feature layer by name in the active map.
        /// </summary>
        /// <param name="layerName"></param>
        /// <returns></returns>
        internal FeatureLayer FindLayer(string layerName)
        {
            // Check there is an input feature layer name.
            if (String.IsNullOrEmpty(layerName))
                return null;

            //IEnumerable<Layer> layers = _activeMap.Layers.Where(layer => layer is FeatureLayer);

            // Finds layers by name and returns a read only list of feature layers.
            IEnumerable<FeatureLayer> layers = _activeMap.FindLayers(layerName, true).OfType<FeatureLayer>();

            while (layers.Any())
            {
                // Get the first feature layer found by name.
                FeatureLayer layer = layers.First();

                // Check the feature layer is in the active map.
                if (layer.Map.Name.Equals(_activeMap.Name, StringComparison.OrdinalIgnoreCase))
                    return layer;
            }

            return null;
        }

        /// <summary>
        /// Remove a layer by name from the active map.
        /// </summary>
        /// <param name="layerName"></param>
        /// <returns></returns>
        public async Task RemoveLayerAsync(string layerName)
        {
            if (String.IsNullOrEmpty(layerName))
                return;

            try
            {
                // Find the layer in the active map.
                FeatureLayer layer = FindLayer(layerName);

                // Remove the layer.
                await RemoveLayerAsync(layer);
            }
            catch
            {
                return;
            }
        }

        /// <summary>
        /// Remove a layer from the active map.
        /// </summary>
        /// <param name="layer"></param>
        /// <returns></returns>
        public async Task RemoveLayerAsync(Layer layer)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    // Remove the layer.
                    if (layer != null)
                        _activeMap.RemoveLayer(layer);
                });
            }
            catch
            {
                // Handle Exception.
            }
        }

        public async Task<int> AddIncrementalNumbersAsync(string outputFeatureClass, string outputLayerName, string labelFieldName, string keyFieldName, int startNumber = 1)
        {
            // Check the obvious.
            if (!await ArcGISFunctions.FeatureClassExistsAsync(outputFeatureClass))
                return -1;

            if (!await FieldExistsAsync(outputLayerName, labelFieldName))
                return -1;

            if (!await FieldIsNumericAsync(outputLayerName, labelFieldName))
                return -1;

            if (!await FieldExistsAsync(outputLayerName, keyFieldName))
                return -1;

            // Get the feature layer.
            FeatureLayer featurelayer = FindLayer(outputLayerName);

            if (featurelayer == null)
                return -1;

            int intStart;
            if (startNumber > 0)
                intStart = startNumber;
            else
                intStart = 1;
            int intMax = intStart - 1;
            int intValue = intMax;

            string keyValue;
            string lastKeyValue = "";

            // Create an edit operation.
            EditOperation editOperation = new();

            try
            {
                await QueuedTask.Run(() =>
                {
                    /// Count the number of features matching the search clause.
                    FeatureClass featureClass = featurelayer.GetFeatureClass();

                    // Get the feature class defintion.
                    using (FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition())
                    {
                        // Get the key field from the definition.
                        ArcGIS.Core.Data.Field keyField = featureClassDefinition.GetFields()
                          .First(x => x.Name.Equals(keyFieldName));

                        // Create a SortDescription for the key field.
                        SortDescription keySortDescription = new(keyField)
                        {
                            CaseSensitivity = CaseSensitivity.Insensitive,
                            SortOrder = ArcGIS.Core.Data.SortOrder.Ascending
                        };

                        // Create a TableSortDescription.
                        TableSortDescription tableSortDescription = new(new List<SortDescription>() { keySortDescription });

                        // Create a cursor of the sorted features.
                        using RowCursor rowCursor = featureClass.Sort(tableSortDescription);
                        while (rowCursor.MoveNext())
                        {
                            // Using the current row.
                            using Row record = rowCursor.Current;

                            // Get the key field value.
                            keyValue = (string)record[keyFieldName];

                            // If the key value is different.
                            if (keyValue != lastKeyValue)
                            {
                                intMax++;
                                intValue = intMax;
                            }

                            editOperation.Modify(record, labelFieldName, intValue);

                            lastKeyValue = keyValue;
                        }
                    }
                });
            }
            catch
            {
                // Handle Exception.
                return 0;
            }

            // Execute the edit operation.
            if (!editOperation.IsEmpty)
            {
                if (!await editOperation.ExecuteAsync())
                {
                    MessageBox.Show(editOperation.ErrorMessage);
                    return -1;
                }
            }

            // Check for unsaved edits.
            if (Project.Current.HasEdits)
            {
                // Save edits.
                await Project.Current.SaveEditsAsync();
            }

            return intMax;
        }

        /// <summary>
        /// Select features in layerName by attributes.
        /// </summary>
        /// <param name="layerName"></param>
        /// <param name="whereClause"></param>
        /// <returns></returns>
        public async Task<bool> SelectLayerByAttributesAsync(string layerName, string whereClause, SelectionCombinationMethod selectionMethod)
        {
            if (String.IsNullOrEmpty(layerName))
                return false;

            try
            {
                // Find the feature layerName by name if it exists. Only search existing layers.
                FeatureLayer featurelayer = FindLayer(layerName);

                if (featurelayer == null)
                    return false;

                // Create a query filter using the where clause.
                QueryFilter queryFilter = new()
                {
                    WhereClause = whereClause
                };

                await QueuedTask.Run(() =>
                {
                    // Select the features matching the search clause.
                    featurelayer.Select(queryFilter, selectionMethod);
                });
            }
            catch
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Select features in layerName by location.
        /// </summary>
        /// <param name="layerName"></param>
        /// <param name="whereClause"></param>
        /// <returns></returns>
        public async Task<bool> SelectLayerByLocationAsync(string targetLayer, string searchLayer, string overlapType = "INTERSECT", string searchDistance = "", string selectionType = "NEW_SELECTION")
        {
            if (String.IsNullOrEmpty(targetLayer) || String.IsNullOrEmpty(searchLayer))
                return false;

            // Make a value array of strings to be passed to the tool.
            IReadOnlyList<string> parameters = Geoprocessing.MakeValueArray(targetLayer, overlapType, searchLayer, searchDistance, selectionType);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread;

            //Geoprocessing.OpenToolDialog("management.SelectLayerByLocation", parameters);  // Useful for debugging.

            //Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.SelectLayerByLocation", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);   // Test this ???

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Clear selected features in layerName.
        /// </summary>
        /// <param name="layerName"></param>
        /// <param name="searchClause"></param>
        /// <returns></returns>
        public async Task<bool> ClearLayerSelectionAsync(string layerName)
        {
            if (String.IsNullOrEmpty(layerName))
                return false;

            try
            {
                // Find the feature layerName by name if it exists. Only search existing layers.
                FeatureLayer featurelayer = FindLayer(layerName);

                if (featurelayer == null)
                    return false;

                await QueuedTask.Run(() =>
                {
                    // Clear the feature selection.
                    featurelayer.ClearSelection();
                });
            }
            catch
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        public async Task<bool> FieldExistsAsync(string layerName, string fieldName)
        {
            if (String.IsNullOrEmpty(layerName))
                return false;

            try
            {
                // Find the feature layerName by name if it exists. Only search existing layers.
                FeatureLayer featurelayer = FindLayer(layerName);   // Need to pass layer name not file name. ???

                if (featurelayer == null)
                    return false;

                IReadOnlyList<ArcGIS.Core.Data.Field> fields = null;
                List<string> fieldList = [];

                bool fldFound = false;

                await QueuedTask.Run(() =>
                {
                    Table table = featurelayer.GetTable();
                    if (table != null)
                    {
                        TableDefinition def = table.GetDefinition();
                        fields = def.GetFields();
                        foreach (ArcGIS.Core.Data.Field fld in fields)
                        {
                            if (fld.Name == fieldName)
                            {
                                fldFound = true;
                                break;
                            }
                        }
                    }
                });

                return fldFound;
            }
            catch
            {
                // Handle Exception.
                return false;
            }
        }

        public async Task<bool> FieldIsNumericAsync(string layerName, string fieldName)
        {
            if (String.IsNullOrEmpty(layerName))
                return false;

            try
            {
                // Find the feature layerName by name if it exists. Only search existing layers.
                FeatureLayer featurelayer = FindLayer(layerName);

                if (featurelayer == null)
                    return false;

                IReadOnlyList<ArcGIS.Core.Data.Field> fields = null;
                List<string> fieldList = [];

                bool fldIsNumeric = false;

                await QueuedTask.Run(() =>
                {
                    Table table = featurelayer.GetTable();
                    if (table != null)
                    {
                        TableDefinition def = table.GetDefinition();
                        fields = def.GetFields();
                        foreach (ArcGIS.Core.Data.Field fld in fields)
                        {
                            if (fld.Name == fieldName)
                            {
                                switch (fld.FieldType)
                                {
                                    case FieldType.SmallInteger:
                                    case FieldType.BigInteger:
                                    case FieldType.Integer:
                                    case FieldType.Single:
                                    case FieldType.Double:
                                        fldIsNumeric = true;
                                        break;
                                    default:
                                        break;
                                }
                                break;
                            }
                        }
                    }
                });

                return fldIsNumeric;
            }
            catch
            {
                // Handle Exception.
                return false;
            }
        }

        public async Task<bool> BufferFeaturesAsync(string inFeatureClass, string outFeatureClass, string bufferDistance, string aggregateFields)
        {
            if (String.IsNullOrEmpty(inFeatureClass))
                return false;

            if (String.IsNullOrEmpty(outFeatureClass))
                return false;

            // Check if all fields in the aggregate fields exist. If not, ignore.
            List<string> aggColumns = [.. aggregateFields.Split(';')];
            aggregateFields = "";
            foreach (string fieldName in aggColumns)
            {
                if (await FieldExistsAsync(inFeatureClass, fieldName))
                {
                    aggregateFields = aggregateFields + fieldName + ";";
                }
            }
            string dissolveOption = "ALL";
            if (aggregateFields != "")
            {
                aggregateFields = aggregateFields.Substring(0, aggregateFields.Length - 1);
                dissolveOption = "LIST";
            }

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inFeatureClass, outFeatureClass, bufferDistance, "FULL", "ROUND", dissolveOption)];
            if (aggregateFields != "")
                parameters.Add(aggregateFields);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.AddOutputsToMap | GPExecuteToolFlags.GPThread | GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("analysis.Buffer", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Buffer", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);   // Test this ???

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        public async Task<bool> ClipFeaturesAsync(string inFeatureClass, string clipFeatureClass, string outFeatureClass, bool addToMap = false)
        {
            if (String.IsNullOrEmpty(inFeatureClass))
                return false;

            if (String.IsNullOrEmpty(clipFeatureClass))
                return false;

            if (String.IsNullOrEmpty(outFeatureClass))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inFeatureClass, clipFeatureClass, outFeatureClass)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.Clip", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Clip", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);   // Test this ???

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        public async Task<bool> IntersectFeaturesAsync(string inFeatureClass, string intersectFeatureClass, string outFeatureClass, bool addToMap = false)
        {
            if (String.IsNullOrEmpty(inFeatureClass))
                return false;

            if (String.IsNullOrEmpty(intersectFeatureClass))
                return false;

            if (String.IsNullOrEmpty(outFeatureClass))
                return false;

            string[] features = ["'" + inFeatureClass + "' #", "'" + intersectFeatureClass + "' #"];
            var featureList = string.Join(";", features);

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(featureList, outFeatureClass)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.Intersect", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Intersect", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);   // Test this ???

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Get the full layer path name for a layer in the map (i.e.
        /// to include any parent group names.
        /// </summary>
        /// <param name="layer"></param>
        /// <returns></returns>
        public string GetLayerPath(Layer layer)
        {
            string layerPath = "";

            // Get the parent for the layer.
            ILayerContainer layerParent = layer.Parent;

            // Loop while the parent is a group layer.
            while (layerParent is GroupLayer)
            {
                Layer grouplayer = (Layer)layerParent;
                layerPath = grouplayer.Name + "/" + layerPath;

                layerParent = grouplayer.Parent;
            }

            return layerPath + layer.Name;
        }

        /// <summary>
        /// Get the full layer path name for a layer in the map (i.e.
        /// to include any parent group names.
        /// </summary>
        /// <param name="layer"></param>
        /// <returns></returns>
        public string GetLayerPath(string layerName)
        {
            if (String.IsNullOrEmpty(layerName))
                return null;

            try
            {
                // Find the layer in the active map.
                FeatureLayer layer = FindLayer(layerName);

                return GetLayerPath(layer);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Returns a simplified feature class shape type: point, line, polygon.
        /// </summary>
        /// <param name="featureLayer"></param>
        /// <returns></returns>
        public string GetFeatureClassType(FeatureLayer featureLayer)
        {
            BasicFeatureLayer basicFeatureLayer = featureLayer as BasicFeatureLayer;
            esriGeometryType shapeType = basicFeatureLayer.ShapeType;

            string classType = "other";
            if (shapeType == esriGeometryType.esriGeometryMultipoint || shapeType == esriGeometryType.esriGeometryPoint)
            {
                classType = "point";
            }
            else if (shapeType == esriGeometryType.esriGeometryRing || shapeType == esriGeometryType.esriGeometryPolygon)
            {
                classType = "polygon";
            }
            else if (shapeType == esriGeometryType.esriGeometryLine || shapeType == esriGeometryType.esriGeometryPolyline ||
                shapeType == esriGeometryType.esriGeometryCircularArc || shapeType == esriGeometryType.esriGeometryEllipticArc ||
                shapeType == esriGeometryType.esriGeometryBezier3Curve || shapeType == esriGeometryType.esriGeometryPath)
            {
                classType = "line";
            }

            return classType;
        }

        /// <summary>
        /// Returns a simplified feature class shape type: point, line, polygon.
        /// </summary>
        /// <param name="featureLayer"></param>
        /// <returns></returns>
        public string GetFeatureClassType(string layerName)
        {
            if (String.IsNullOrEmpty(layerName))
                return null;

            try
            {
                // Find the layer in the active map.
                FeatureLayer layer = FindLayer(layerName);

                return GetFeatureClassType(layer);
            }
            catch
            {
                return null;
            }
        }

        #endregion Layers

        #region Group Layers

        /// <summary>
        /// Find a group layer by name in the active map.
        /// </summary>
        /// <param name="layerName"></param>
        /// <returns></returns>
        internal GroupLayer FindGroupLayer(string groupLayerName)
        {
            // Check there is an input groupLayer name.
            if (String.IsNullOrEmpty(groupLayerName))
                return null;

            // Finds group layers by name and returns a read only list of group layers.
            IEnumerable<GroupLayer> groupLayers = _activeMap.FindLayers(groupLayerName).OfType<GroupLayer>();

            while (groupLayers.Any())
            {
                // Get the first group layer found by name.
                GroupLayer groupLayer = groupLayers.First();

                // Check the group layer is in the active map.
                if (groupLayer.Map.Name.Equals(_activeMap.Name, StringComparison.OrdinalIgnoreCase))
                    return groupLayer;
            }

            return null;
        }

        /// <summary>
        /// Move a layer into a group layer (creating the group layer if
        /// it doesn't already exist.
        /// </summary>
        /// <param name="groupLayerName"></param>
        /// <param name="layer"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        public async Task<bool> MoveToGroupLayerAsync(Layer layer, string groupLayerName, int position)
        {
            // Check there is an input groupLayer name.
            if (String.IsNullOrEmpty(groupLayerName))
                return false;

            // Check if there is an input layer.
            if (layer == null)
                return false;

            // Does the group layer exist?
            GroupLayer groupLayer = FindGroupLayer(groupLayerName);
            if (groupLayer == null)
            {
                // Add the group layer to the map.
                try
                {
                    await QueuedTask.Run(() =>
                    {
                        groupLayer = LayerFactory.Instance.CreateGroupLayer(_activeMap, 0, groupLayerName);
                    });
                }
                catch
                {
                    // Handle Exception.
                    return false;
                }
            }

            // Move the layer into the group.
            try
            {
                await QueuedTask.Run(() =>
                {
                    // Move the layer into the group.
                    _activeMap.MoveLayer(layer, groupLayer, position);

                    // Expand the group.
                    groupLayer.SetExpanded(true);
                });
            }
            catch
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Remove a group layer if it is empty.
        /// </summary>
        /// <param name="groupLayerName"></param>
        /// <returns></returns>
        public async Task<bool> RemoveGroupLayerAsync(string groupLayerName)
        {
            // Check there is an input groupLayer name.
            if (String.IsNullOrEmpty(groupLayerName))
                return false;

            // Does the group layer exist?
            GroupLayer groupLayer = FindGroupLayer(groupLayerName);
            if (groupLayer == null)
                return false;

            // Count the layers in the group.
            if (groupLayer.Layers.Count != 0)
                return true;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Remove the group layer.
                    _activeMap.RemoveLayer(groupLayer);
                });
            }
            catch
            {
                // Handle Exception.
                return false;
            }

            return true;
        }
        #endregion Group Layers

        #region Tables

        /// <summary>
        /// Find a table by name in the active map.
        /// </summary>
        /// <param name="tableName"></param>
        /// <returns></returns>
        internal StandaloneTable FindTable(string tableName)
        {
            // Check there is an input table name.
            if (String.IsNullOrEmpty(tableName))
                return null;

            // Finds tables by name and returns a read only list of standalone tables.
            IEnumerable<StandaloneTable> tables = _activeMap.FindStandaloneTables(tableName).OfType<StandaloneTable>();

            while (tables.Any())
            {
                // Get the first table found by name.
                StandaloneTable table = tables.First();

                // Check the table is in the active map.
                if (table.Map.Name.Equals(_activeMap.Name, StringComparison.OrdinalIgnoreCase))
                    return table;
            }

            return null;
        }

        public bool RemoveTable(string tableName)
        {
            try
            {
                // Find the table in the active map.
                StandaloneTable table = FindTable(tableName);

                if (table != null)
                {
                    // Remove the table.
                    _activeMap.RemoveStandaloneTable(table);
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        #endregion Tables

        #region Symbology

        /// <summary>
        /// Apply symbology to a layer by name using a lyrx file.
        /// </summary>
        /// <param name="layerName"></param>
        /// <param name="layerFile"></param>
        /// <returns></returns>
        public async Task<bool> ApplySymbologyFromLayerFileAsync(string layerName, string layerFile)
        {
            // Check there is an input layer name.
            if (String.IsNullOrEmpty(layerName))
                return false;

            // Check the lyrx file exists.
            if (!FileFunctions.FileExists(layerFile))
                return false;

            // Find the layer in the active map.
            FeatureLayer featureLayer = FindLayer(layerName);

            if (featureLayer != null)
            {
                // Apply the layer file symbology to the feature layer.
                try
                {
                    await QueuedTask.Run(() =>
                    {
                        // Get the Layer Document from the lyrx file.
                        LayerDocument lyrDocFromLyrxFile = new(layerFile);

                        CIMLayerDocument cimLyrDoc = lyrDocFromLyrxFile.GetCIMLayerDocument();

                        // Get the renderer from the layer file.
                        //CIMSimpleRenderer rendererFromLayerFile = ((CIMFeatureLayer)cimLyrDoc.LayerDefinitions[0]).Renderer as CIMSimpleRenderer;
                        var rendererFromLayerFile = ((CIMFeatureLayer)cimLyrDoc.LayerDefinitions[0]).Renderer;

                        // Apply the renderer to the feature layer.
                        if (featureLayer.CanSetRenderer(rendererFromLayerFile))
                            featureLayer.SetRenderer(rendererFromLayerFile);
                    });
                }
                catch
                {
                    // Handle Exception.
                    return false;
                }
            }

            return true;
        }

        #endregion Symbology
    }

    /// <summary>
    /// This helper class provides ArcGIS Pro feature class and layer functions.
    /// </summary>
    internal static class ArcGISFunctions
    {
        #region Feature Class

        /// <summary>
        /// Check if the feature class exists in the file path.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static async Task<bool> FeatureClassExistsAsync(string filePath, string fileName)
        {
            if (fileName.Substring(fileName.Length - 4, 1) == ".")
            {
                // It's a file.
                if (FileFunctions.FileExists(filePath + @"\" + fileName))
                    return true;
                else
                    return false;
            }
            else if (filePath.Substring(filePath.Length - 3, 3) == "sde")
            {
                // It's an SDE class
                // Not handled. We know the layer exists.
                return true;
            }
            else // it is a geodatabase class.
            {
                try
                {
                    bool exists = await FeatureClassExistsGDBAsync(filePath, fileName);

                    return exists;
                }
                catch
                {
                    // GetDefinition throws an exception if the definition doesn't exist
                    return false;
                }
            }
        }

        /// <summary>
        /// Check if the feature class exists.
        /// </summary>
        /// <param name="fullPath"></param>
        /// <returns></returns>
        public static async Task<bool> FeatureClassExistsAsync(string fullPath)
        {
            return await FeatureClassExistsAsync(FileFunctions.GetDirectoryName(fullPath), FileFunctions.GetFileName(fullPath));
        }


        /// <summary>
        /// Delete a feature class from a geodatabase.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        public static void DeleteFeatureClass(string filePath, string fileName)
        {
            try
            {
                // Open the file geodatabase. This will open the geodatabase if the folder exists and contains a valid geodatabase.
                using Geodatabase geodatabase = new(new FileGeodatabaseConnectionPath(new Uri(filePath)));

                DeleteFeatureClass(geodatabase, fileName);
            }
            catch { }
        }

        /// <summary>
        /// Delete a feature class from a geodatabase.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="fileName"></param>
        public static void DeleteFeatureClass(Geodatabase geodatabase, string fileName)
        {
            try
            {
                // Create a SchemaBuilder object
                SchemaBuilder schemaBuilder = new(geodatabase);

                // Create a FeatureClassDescription object.
                using FeatureClassDefinition featureClassDefinition = geodatabase.GetDefinition<FeatureClassDefinition>(fileName);

                // Create a FeatureClassDescription object
                FeatureClassDescription featureClassDescription = new(featureClassDefinition);

                // Add the deletion for the feature class to the list of DDL tasks
                schemaBuilder.Delete(featureClassDescription);

                // Execute the DDL
                bool success = schemaBuilder.Build();
            }
            catch { }
        }

        /// <summary>
        /// Delete a feature class.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        public static async Task<bool> DeleteFeatureClassAsync(string filePath, string fileName)
        {
            string featureClass = filePath + @"\" + fileName;

            return await DeleteFeatureClassAsync(featureClass);
        }

        /// <summary>
        /// Copy the input feature class to the output feature class.
        /// </summary>
        /// <param name="inFeatureClass"></param>
        /// <param name="outFeatureClass"></param>
        /// <param name="Messages"></param>
        /// <returns></returns>
        public static async Task<bool> DeleteFeatureClassAsync(string fileName)
        {
            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(fileName);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread;

            //Geoprocessing.OpenToolDialog("management.Delete", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.Delete", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);   // Test this ???

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        public static async Task<bool> AddFieldToFeatureClassAsync(string inTable, string fieldName, string fieldType,
            long fieldPrecision = -1, long fieldScale = -1, long fieldLength = -1,
            string fieldAlias = null, bool fieldIsNullable = true, bool fieldIsRequred = false, string fieldDomain = null)
        {
            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, fieldName, fieldType,
                fieldPrecision > 0 ? fieldPrecision : null, fieldScale > 0 ? fieldScale : null, fieldLength > 0 ? fieldLength : null,
                fieldAlias ?? null, fieldIsNullable ? "NULLABLE" : "NON_NULLABLE",
                fieldIsRequred ? "REQUIRED" : "NON_REQUIRED", fieldDomain);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread;

            //Geoprocessing.OpenToolDialog("management.AddField", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.AddField", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);   // Test this ???

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Count the features in a layer using a search where clause.
        /// </summary>
        /// <param name="layer"></param>
        /// <param name="whereClause"></param>
        /// <returns></returns>
        public static async Task<long> CountFeaturesAsync(FeatureLayer layer, string whereClause)
        {
            long featureCount = 0;

            if (layer == null)
                return featureCount;

            try
            {
                // Create a query filter using the where clause.
                QueryFilter queryFilter = new()
                {
                    WhereClause = whereClause
                };

                featureCount = await QueuedTask.Run(() =>
                {
                    /// Count the number of features matching the search clause.
                    FeatureClass featureClass = layer.GetFeatureClass();

                    return featureClass.GetCount(queryFilter);
                });
            }
            catch
            {
                // Handle Exception.
                return 0;
            }

            return featureCount;
        }

        #endregion Feature Class

        #region Geodatabase

        public static Geodatabase CreateFileGeodatabase(string fullPath)
        {
            // Create a FileGeodatabaseConnectionPath with the name of the file geodatabase you wish to create
            FileGeodatabaseConnectionPath fileGeodatabaseConnectionPath = new(new Uri(fullPath));

            // Create and use the file geodatabase
            Geodatabase geodatabase = SchemaBuilder.CreateGeodatabase(fileGeodatabaseConnectionPath);

            return geodatabase;
        }

        /// <summary>
        /// Check if the feature class exists in a geodatabase.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static async Task<bool> FeatureClassExistsGDBAsync(string filePath, string fileName)
        {
            bool exists = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Open the file geodatabase. This will open the geodatabase if the folder exists and contains a valid geodatabase.
                    using Geodatabase geodatabase = new(new FileGeodatabaseConnectionPath(new Uri(filePath)));

                    // Create a FeatureClassDefinition object.
                    using FeatureClassDefinition featureClassDefinition = geodatabase.GetDefinition<FeatureClassDefinition>(fileName);

                    if (featureClassDefinition != null)
                        exists = true;
                });
            }
            catch (GeodatabaseNotFoundOrOpenedException)
            {
                // Handle Exception.
                return false;
            }

            return exists;
        }

        /// <summary>
        /// Check if the layer exists in a geodatabase.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static async Task<bool> TableExistsGDBAsync(string filePath, string fileName)
        {
            bool exists = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Open the file geodatabase. This will open the geodatabase if the folder exists and contains a valid geodatabase.
                    using Geodatabase geodatabase = new(new FileGeodatabaseConnectionPath(new Uri(filePath)));

                    // Create a TableDefinition object.
                    using TableDefinition tableDefinition = geodatabase.GetDefinition<TableDefinition>(fileName);

                    if (tableDefinition != null)
                        exists = true;
                });
            }
            catch (GeodatabaseNotFoundOrOpenedException)
            {
                // Handle Exception.
                return false;
            }

            return exists;
        }
        #endregion Geodatabase

        #region Table

        /// <summary>
        /// Check if the feature class exists in the file path.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static async Task<bool> TableExistsAsync(string filePath, string fileName)
        {
            if (fileName.Substring(fileName.Length - 4, 1) == ".")
            {
                // It's a file.
                if (FileFunctions.FileExists(filePath + @"\" + fileName))
                    return true;
                else
                    return false;
            }
            else if (filePath.Substring(filePath.Length - 3, 3) == "sde")
            {
                // It's an SDE class
                // Not handled. We know the layer exists.
                return true;
            }
            else // it is a geodatabase class.
            {
                try
                {
                    bool exists = await TableExistsGDBAsync(filePath, fileName);

                    return exists;
                }
                catch
                {
                    // GetDefinition throws an exception if the definition doesn't exist
                    return false;
                }
            }
        }

        /// <summary>
        /// Check if the feature class exists.
        /// </summary>
        /// <param name="fullPath"></param>
        /// <returns></returns>
        public static async Task<bool> TableExistsAsync(string fullPath)
        {
            return await TableExistsAsync(FileFunctions.GetDirectoryName(fullPath), FileFunctions.GetFileName(fullPath));
        }

        /// <summary>
        /// Check the layer exists in the file path.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static bool TableExists(string filePath, string fileName)
        {
            if (fileName.Substring(fileName.Length - 4, 1) == ".")
            {
                // It's a file.
                if (FileFunctions.FileExists(filePath + @"\" + fileName))
                    return true;
                else
                    return false;
            }
            else if (filePath.Substring(filePath.Length - 3, 3) == "sde")
            {
                // It's an SDE class
                // Not handled. We know the layer exists.
                return true;
            }
            else // it is a geodatabase class.
            {
                //IWorkspaceFactory pWSF = GetWorkspaceFactory(filePath);
                //IWorkspace2 pWS = (IWorkspace2)pWSF.OpenFromFile(filePath, 0);
                //if (pWS.get_NameExists(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTTable, tableName))
                //    return true;
                //else
                //    return false;
                return false;
            }
        }

        /// <summary>
        /// Check if the layer exists.
        /// </summary>
        /// <param name="fullPath"></param>
        /// <returns></returns>
        public static bool TableExists(string fullPath)
        {
            return TableExists(FileFunctions.GetDirectoryName(fullPath), FileFunctions.GetFileName(fullPath));
        }

        /// <summary>
        /// Delete a table from a geodatabase.
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        public static void DeleteTable(string filePath, string fileName)
        {
            try
            {
                // Open the file geodatabase. This will open the geodatabase if the folder exists and contains a valid geodatabase.
                using Geodatabase geodatabase = new(new FileGeodatabaseConnectionPath(new Uri(filePath)));

                DeleteTable(geodatabase, fileName);
            }
            catch { }
        }

        /// <summary>
        /// Delete a table from a geodatabase.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="tableName"></param>
        public static void DeleteTable(Geodatabase geodatabase, string tableName)
        {
            try
            {
                // Create a SchemaBuilder object
                SchemaBuilder schemaBuilder = new(geodatabase);

                // Create a FeatureClassDescription object.
                using TableDefinition tableDefinition = geodatabase.GetDefinition<TableDefinition>(tableName);

                // Create a FeatureClassDescription object
                TableDescription tableDescription = new(tableDefinition);

                // Add the deletion for the feature class to the list of DDL tasks
                schemaBuilder.Delete(tableDescription);

                // Execute the DDL
                bool success = schemaBuilder.Build();
            }
            catch { }
        }

        #endregion Table

        #region Outputs

        /// <summary>
        /// Prompt the user to specify an output file in the required format.
        /// </summary>
        /// <param name="fileType"></param>
        /// <param name="initialDirectory"></param>
        /// <returns></returns>
        public static string GetOutputFileName(string fileType, string initialDirectory = @"C:\")
        {
            BrowseProjectFilter bf;

            //string saveItemDlg;
            switch (fileType)
            {
                case "Geodatabase FC":
                    bf = BrowseProjectFilter.GetFilter("esri_browseDialogFilters_geodatabaseItems_featureClasses");
                    break;

                case "Geodatabase Table":
                    bf = BrowseProjectFilter.GetFilter("esri_browseDialogFilters_geodatabaseItems_tables");
                    break;

                case "Shapefile":
                    bf = BrowseProjectFilter.GetFilter("esri_browseDialogFilters_shapefiles");
                    break;

                case "CSV file (comma delimited)":
                    bf = BrowseProjectFilter.GetFilter("esri_browseDialogFilters_textFiles_csv");
                    break;

                case "Text file (tab delimited)":
                    bf = BrowseProjectFilter.GetFilter("esri_browseDialogFilters_textFiles_txt");
                    break;

                default:
                    bf = BrowseProjectFilter.GetFilter("esri_browseDialogFilters_all");
                    break;
            }

            // Display the saveItemDlg in an Open Item dialog.
            SaveItemDialog saveItemDlg = new()
            {
                Title = "Save Output As...",
                InitialLocation = initialDirectory,
                //AlwaysUseInitialLocation = true,
                //Filter = ItemFilters.Files_All,
                OverwritePrompt = false,    // This will be done later.
                BrowseFilter = bf
            };

            bool? ok = saveItemDlg.ShowDialog();

            string strOutFile = null;
            if (ok.HasValue)
                strOutFile = saveItemDlg.FilePath;

            return strOutFile; // Null if user pressed exit
        }

        #endregion Outputs

        #region CopyFeatures

        /// <summary>
        /// Copy the input feature class to the output feature class.
        /// </summary>
        /// <param name="inFeatureClass"></param>
        /// <param name="outFeatureClass"></param>
        /// <param name="Messages"></param>
        /// <returns></returns>
        public static async Task<bool> CopyFeaturesAsync(string inFeatureClass, string outFeatureClass, bool addToMap = false)
        {
            if (String.IsNullOrEmpty(inFeatureClass))
                return false;

            if (String.IsNullOrEmpty(outFeatureClass))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inFeatureClass, outFeatureClass);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("management.CopyFeatures", parameters);  // Useful for debugging.

                // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.CopyFeatures", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);   // Test this ???

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Copy the input dataset name to the output feature class.
        /// </summary>
        /// <param name="InWorkspace"></param>
        /// <param name="InDatasetName"></param>
        /// <param name="OutFeatureClass"></param>
        /// <param name="Messages"></param>
        /// <returns></returns>
        public static async Task<bool> CopyFeaturesAsync(string InWorkspace, string InDatasetName, string OutFeatureClass)
        {
            string inFeatureClass = InWorkspace + @"\" + InDatasetName;

            return await CopyFeaturesAsync(inFeatureClass, OutFeatureClass);
        }

        /// <summary>
        /// Copy the input dataset to the output dataset.
        /// </summary>
        /// <param name="InWorkspace"></param>
        /// <param name="InDatasetName"></param>
        /// <param name="OutWorkspace"></param>
        /// <param name="OutDatasetName"></param>
        /// <param name="Messages"></param>
        /// <returns></returns>
        public static async Task<bool> CopyFeaturesAsync(string InWorkspace, string InDatasetName, string OutWorkspace, string OutDatasetName)
        {
            string inFeatureClass = InWorkspace + @"\" + InDatasetName;
            string outFeatureClass = OutWorkspace + @"\" + OutDatasetName;

            return await CopyFeaturesAsync(inFeatureClass, outFeatureClass);
        }

        #endregion CopyFeatures

        #region Export Features

        /// <summary>
        /// Export the input table to the output table.
        /// </summary>
        /// <param name="InTable"></param>
        /// <param name="OutFile"></param>
        /// <param name="Messages"></param>
        /// <returns></returns>
        public static async Task<bool> ExportFeaturesAsync(string inTable, string outTable)
        {
            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, outTable);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread;

            //Geoprocessing.OpenToolDialog("conversion.ExportTable", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("conversion.ExportTable", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);   // Test this ???

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        #endregion Export Features

        #region Copy Table

        /// <summary>
        /// Copy the input table to the output table.
        /// </summary>
        /// <param name="inTable"></param>
        /// <param name="outTable"></param>
        /// <param name="Messages"></param>
        /// <returns></returns>
        public static async Task<bool> CopyTableAsync(string inTable, string outTable)
        {
            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, outTable);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread;

            //Geoprocessing.OpenToolDialog("management.Copy", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.Copy", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);   // Test this ???

                    var messages = gp_result.Messages;
                    var errMessages = gp_result.ErrorMessages;
                    return false;
                }
            }
            catch (Exception)
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Copy the input dataset name to the output table.
        /// </summary>
        /// <param name="InWorkspace"></param>
        /// <param name="InDatasetName"></param>
        /// <param name="OutTable"></param>
        /// <param name="Messages"></param>
        /// <returns></returns>
        public static async Task<bool> CopyTableAsync(string InWorkspace, string InDatasetName, string OutTable)
        {
            string inTable = InWorkspace + @"\" + InDatasetName;
            return await CopyTableAsync(inTable, OutTable);
        }

        /// <summary>
        /// Copy the input dataset to the output dataset.
        /// </summary>
        /// <param name="InWorkspace"></param>
        /// <param name="InDatasetName"></param>
        /// <param name="OutWorkspace"></param>
        /// <param name="OutDatasetName"></param>
        /// <param name="Messages"></param>
        /// <returns></returns>
        public static async Task<bool> CopyTableAsync(string InWorkspace, string InDatasetName, string OutWorkspace, string OutDatasetName)
        {
            string inTable = InWorkspace + @"\" + InDatasetName;
            string outTable = OutWorkspace + @"\" + OutDatasetName;
            return await CopyTableAsync(inTable, outTable);
        }

        #endregion Copy Table
    }
}