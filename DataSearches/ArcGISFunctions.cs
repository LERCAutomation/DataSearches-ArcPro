﻿// The Data tools are a suite of ArcGIS Pro addins used to extract
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
using ArcGIS.Desktop.Catalog;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Editing;
using ArcGIS.Desktop.Editing.Attributes;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using QueryFilter = ArcGIS.Core.Data.QueryFilter;

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

        public void PauseDrawing(bool pause)
        {
            _activeMapView.DrawingPaused = pause;
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
        public async Task<bool> AddLayerToMap(string url)
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
                return false;
            }

            return true;
        }

        /// <summary>
        /// Add a standalone layer to the active map.
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public async Task<bool> AddTableToMap(string url)
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
                return false;
            }

            return true;
        }

        public async Task<bool> ZoomToLayerAsync(string layerName, double ratio = 1)
        {
            // Check if the layer is already loaded.
            Layer findLayer = FindLayer(layerName);

            // If the layer is not loaded, add it.
            if (findLayer == null)
                return false;

            try
            {
                // Zoom to the layer extent.
                await _activeMapView.ZoomToAsync(findLayer, false);

                // Get the camera for the active view, adjust the scale and zoom to the new camera position.
                var camera = _activeMapView.Camera;
                camera.Scale *= ratio;
                await _activeMapView.ZoomToAsync(camera);
            }
            catch
            {
                // Handle exception.
                return false;
            }

            return true;
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
        public async Task<bool> RemoveLayerAsync(string layerName)
        {
            if (String.IsNullOrEmpty(layerName))
                return false;

            try
            {
                // Find the layer in the active map.
                FeatureLayer layer = FindLayer(layerName);

                // Remove the layer.
                if (layer != null)
                    return await RemoveLayerAsync(layer);
            }
            catch
            {
                // Handle exception.
                return false;
            }

            return true;
        }

        /// <summary>
        /// Remove a layer from the active map.
        /// </summary>
        /// <param name="layer"></param>
        /// <returns></returns>
        public async Task<bool> RemoveLayerAsync(Layer layer)
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
                return false;
            }

            return true;
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
            FeatureLayer outputFeaturelayer = FindLayer(outputLayerName);
            if (outputFeaturelayer == null)
                return -1;

            int labelMax;
            if (startNumber > 1)
                labelMax = startNumber - 1;
            else
                labelMax = 0;
            int labelVal = labelMax;

            string keyValue;
            string lastKeyValue = "";

            // Create an edit operation.
            EditOperation editOperation = new();

            try
            {
                await QueuedTask.Run(() =>
                {
                    /// Get the feature class for the output feature layer.
                    FeatureClass featureClass = outputFeaturelayer.GetFeatureClass();

                    // Get the feature class defintion.
                    using FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();

                    // Get the key field from the feature class definition.
                    ArcGIS.Core.Data.Field keyField = featureClassDefinition.GetFields()
                      .First(x => x.Name.Equals(keyFieldName));

                    // Create a SortDescription for the key field.
                    ArcGIS.Core.Data.SortDescription keySortDescription = new(keyField)
                    {
                        CaseSensitivity = CaseSensitivity.Insensitive,
                        SortOrder = ArcGIS.Core.Data.SortOrder.Ascending
                    };

                    // Create a TableSortDescription.
                    TableSortDescription tableSortDescription = new([keySortDescription]);

                    // Create a cursor of the sorted features.
                    using RowCursor rowCursor = featureClass.Sort(tableSortDescription);
                    while (rowCursor.MoveNext())
                    {
                        // Using the current row.
                        using Row record = rowCursor.Current;

                        // Get the key field value.
                        keyValue = record[keyFieldName].ToString();

                        // If the key value is different.
                        if (keyValue != lastKeyValue)
                        {
                            labelMax++;
                            labelVal = labelMax;
                        }

                        //editOperation.Modify(record, labelFieldName, labelVal); //TODO: Temporarily commented out.

                        lastKeyValue = keyValue;
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
                    //MessageBox.Show(editOperation.ErrorMessage);
                    return -1;
                }
            }

            // Check for unsaved edits.
            if (Project.Current.HasEdits)
            {
                // Save edits.
                await Project.Current.SaveEditsAsync();
            }

            return labelMax;
        }

        public async Task<bool> UpdateFeaturesAsync(string layerName, string siteColumn, string siteName, string orgColumn, string orgName, string radiusColumn, string radiusText)
        {
            if (String.IsNullOrEmpty(layerName))
                return false;

            if (String.IsNullOrEmpty(siteColumn) && String.IsNullOrEmpty(orgColumn) && String.IsNullOrEmpty(radiusColumn))
                return false;

            if (!string.IsNullOrEmpty(siteColumn) && !await FieldExistsAsync(layerName, siteColumn))
                return false;

            if (!string.IsNullOrEmpty(orgColumn) && !await FieldExistsAsync(layerName, orgColumn))
                return false;

            if (!string.IsNullOrEmpty(radiusColumn) && !await FieldExistsAsync(layerName, radiusColumn))
                return false;

            // Get the feature layer.
            FeatureLayer featurelayer = FindLayer(layerName);

            if (featurelayer == null)
                return false;

            // Create an edit operation.
            EditOperation editOperation = new();

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Get the oids for the selected features.
                    var gsSelection = featurelayer.GetSelection();
                    IReadOnlyList<long> selectedOIDs = gsSelection.GetObjectIDs();

                    // Update the attributes of the selected features.
                    var insp = new Inspector();
                    insp.Load(featurelayer, selectedOIDs);

                    if (!string.IsNullOrEmpty(siteColumn))
                    {
                        // Double check that attribute exists.
                        ArcGIS.Desktop.Editing.Attributes.Attribute att = insp.FirstOrDefault(a => a.FieldName == siteColumn);
                        if (att != null)
                            insp[siteColumn] = siteName;
                    }

                    if (!string.IsNullOrEmpty(orgColumn))
                    {
                        // Double check that attribute exists.
                        ArcGIS.Desktop.Editing.Attributes.Attribute att = insp.FirstOrDefault(a => a.FieldName == orgColumn);
                        if (att != null)
                            insp[orgColumn] = orgName;
                    }

                    if (!string.IsNullOrEmpty(radiusColumn))
                    {
                        // Double check that attribute exists.
                        ArcGIS.Desktop.Editing.Attributes.Attribute att = insp.FirstOrDefault(a => a.FieldName == radiusColumn);
                        if (att != null)
                            insp[radiusColumn] = radiusText;
                    }

                    editOperation.Modify(insp);
                });
            }
            catch
            {
                // Handle Exception.
                return false;
            }

            // Execute the edit operation.
            if (!editOperation.IsEmpty)
            {
                if (!await editOperation.ExecuteAsync())
                {
                    MessageBox.Show(editOperation.ErrorMessage);
                    return false;
                }
            }

            // Check for unsaved edits.
            if (Project.Current.HasEdits)
            {
                // Save edits.
                return await Project.Current.SaveEditsAsync();
            }

            return true;
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

        public async Task<IReadOnlyList<ArcGIS.Core.Data.Field>> GetFCFieldsAsync(string layerPath)
        {
            if (String.IsNullOrEmpty(layerPath))
                return null;

            try
            {
                // Find the feature layer by name if it exists. Only search existing layers.
                FeatureLayer featurelayer = FindLayer(layerPath);

                if (featurelayer == null)
                    return null;

                IReadOnlyList<ArcGIS.Core.Data.Field> fields = null;
                List<string> fieldList = [];

                await QueuedTask.Run(() =>
                {
                    // Get the underlying feature class as a table.
                    ArcGIS.Core.Data.Table table = featurelayer.GetTable();
                    if (table != null)
                    {
                        // Get the table definition of the table.
                        TableDefinition tableDef = table.GetDefinition();

                        // Get the fields in the table.
                        fields = tableDef.GetFields();
                    }

                    table.Dispose();
                });

                return fields;
            }
            catch
            {
                // Handle Exception.
                return null;
            }
        }

        public async Task<IReadOnlyList<ArcGIS.Core.Data.Field>> GetTableFieldsAsync(string layerPath)
        {
            if (String.IsNullOrEmpty(layerPath))
                return null;

            try
            {
                // Find the table by name if it exists. Only search existing layers.
                StandaloneTable inputTable = FindTable(layerPath);

                if (inputTable == null)
                    return null;

                IReadOnlyList<ArcGIS.Core.Data.Field> fields = null;
                List<string> fieldList = [];

                await QueuedTask.Run(() =>
                {
                    // Get the underlying table.
                    ArcGIS.Core.Data.Table table = inputTable.GetTable();
                    if (table != null)
                    {
                        // Get the table definition of the table.
                        TableDefinition tableDef = table.GetDefinition();

                        // Get the fields in the table.
                        fields = tableDef.GetFields();
                    }

                    table.Dispose();
                });

                return fields;
            }
            catch
            {
                // Handle Exception.
                return null;
            }
        }

        public bool FieldExists(IReadOnlyList<ArcGIS.Core.Data.Field> fields, string fieldName)
        {
            bool fldFound = false;

            foreach (ArcGIS.Core.Data.Field fld in fields)
            {
                if (fld.Name == fieldName || fld.AliasName == fieldName)
                {
                    fldFound = true;
                    break;
                }
            }

            return fldFound;
        }

        public async Task<bool> FieldExistsAsync(string layerPath, string fieldName)
        {
            if (String.IsNullOrEmpty(layerPath))
                return false;

            try
            {
                // Find the feature layer by name if it exists. Only search existing layers.
                FeatureLayer featurelayer = FindLayer(layerPath);

                if (featurelayer == null)
                    return false;

                IReadOnlyList<ArcGIS.Core.Data.Field> fields = null;

                bool fldFound = false;

                await QueuedTask.Run(() =>
                {
                    // Get the underlying feature class as a table.
                    ArcGIS.Core.Data.Table table = featurelayer.GetTable();
                    if (table != null)
                    {
                        // Get the table definition of the table.
                        TableDefinition tableDef = table.GetDefinition();

                        // Get the fields in the table.
                        fields = tableDef.GetFields();

                        // Loop through all fields looking for a name match.
                        foreach (ArcGIS.Core.Data.Field fld in fields)
                        {
                            if (fld.Name == fieldName || fld.AliasName == fieldName)
                            {
                                fldFound = true;
                                break;
                            }
                        }
                    }

                    table.Dispose();
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
                    // Get the underlying feature class as a table.
                    ArcGIS.Core.Data.Table table = featurelayer.GetTable();
                    if (table != null)
                    {
                        // Get the table definition of the table.
                        TableDefinition tableDef = table.GetDefinition();

                        // Get the fields in the table.
                        fields = tableDef.GetFields();

                        // Loop through all fields looking for a name match.
                        foreach (ArcGIS.Core.Data.Field fld in fields)
                        {
                            if (fld.Name == fieldName || fld.AliasName == fieldName)
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

        /// <summary>
        /// Get the full layer path name for a layer in the map (i.e.
        /// to include any parent group names.
        /// </summary>
        /// <param name="layer"></param>
        /// <returns></returns>
        public string GetLayerPath(Layer layer)
        {
            if (layer == null)
                return null;

            string layerPath = "";

            try
            {
                // Get the parent for the layer.
                ILayerContainer layerParent = layer.Parent;

                // Loop while the parent is a group layer.
                while (layerParent is GroupLayer)
                {
                    // Get the parent layer.
                    Layer grouplayer = (Layer)layerParent;

                    // Append the parent name to the full layer path.
                    layerPath = grouplayer.Name + "/" + layerPath;

                    // Get the parent for the layer.
                    layerParent = grouplayer.Parent;
                }
            }
            catch
            {
                // Handle Exception.
                return null;
            }

            // Append the layer name to it's full path.
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

                if (layer == null)
                    return null;

                return GetLayerPath(layer);
            }
            catch
            {
                // Handle Exception.
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
            try
            {
                BasicFeatureLayer basicFeatureLayer = featureLayer as BasicFeatureLayer;
                esriGeometryType shapeType = basicFeatureLayer.ShapeType;

                return shapeType switch
                {
                    esriGeometryType.esriGeometryPoint => "point",
                    esriGeometryType.esriGeometryMultipoint => "point",
                    esriGeometryType.esriGeometryPolygon => "polygon",
                    esriGeometryType.esriGeometryRing => "polygon",
                    esriGeometryType.esriGeometryLine => "line",
                    esriGeometryType.esriGeometryPolyline => "line",
                    esriGeometryType.esriGeometryCircularArc => "line",
                    esriGeometryType.esriGeometryEllipticArc => "line",
                    esriGeometryType.esriGeometryBezier3Curve => "line",
                    esriGeometryType.esriGeometryPath => "line",
                    _ => "other",
                };
            }
            catch (Exception)
            {
                // Handle the exception.
                return null;
            }
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

                if (layer == null)
                    return null;

                return GetFeatureClassType(layer);
            }
            catch
            {
                // Handle Exception.
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
        internal GroupLayer FindGroupLayer(string layerName)
        {
            // Check there is an input groupLayer name.
            if (String.IsNullOrEmpty(layerName))
                return null;

            try
            {
                // Finds group layers by name and returns a read only list of group layers.
                IEnumerable<GroupLayer> groupLayers = _activeMap.FindLayers(layerName).OfType<GroupLayer>();

                while (groupLayers.Any())
                {
                    // Get the first group layer found by name.
                    GroupLayer groupLayer = groupLayers.First();

                    // Check the group layer is in the active map.
                    if (groupLayer.Map.Name.Equals(_activeMap.Name, StringComparison.OrdinalIgnoreCase))
                        return groupLayer;
                }
            }
            catch
            {
                // Handle exception.
                return null;
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
            // Check there is an input groupLayerName name.
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

            try
            {
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
            }
            catch
            {
                // Handle exception.
                return null;
            }

            return null;
        }

        public async Task<bool> RemoveTableAsync(string tableName)
        {
            try
            {
                // Find the table in the active map.
                StandaloneTable table = FindTable(tableName);

                if (table != null)
                {
                    // Remove the table.
                    await RemoveTableAsync(table);
                }

                return true;
            }
            catch
            {
                // Handle exception.
                return false;
            }
        }

        /// <summary>
        /// Remove a table from the active map.
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public async Task<bool> RemoveTableAsync(StandaloneTable table)
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    // Remove the table.
                    if (table != null)
                        _activeMap.RemoveStandaloneTable(table);
                });
            }
            catch
            {
                // Handle Exception.
                return false;
            }

            return true;
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

        public async Task<bool> LabelLayerAsync(string layerName, string labelColumn, string labelFont = "Arial", double labelSize = 10, string labelStyle = "Normal",
                            int labelRed = 0, int labelGreen = 0, int labelBlue = 0, string labelOverlap = "OnePerShape", bool displayLabels = true)
        {
            // Check there is an input layer.
            if (String.IsNullOrEmpty(layerName))
                return false;

            // Check there is a label columns to set.
            if (String.IsNullOrEmpty(labelColumn))
                return false;

            // Get the input feature layer.
            FeatureLayer featurelayer = FindLayer(layerName);

            if (featurelayer == null)
                return false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    CIMColor textColor = ColorFactory.Instance.CreateRGBColor(labelRed, labelGreen, labelBlue);

                    CIMTextSymbol textSymbol = SymbolFactory.Instance.ConstructTextSymbol(textColor, labelSize, labelFont, labelStyle);

                    // Get the layer definition.
                    CIMFeatureLayer lyrDefn = featurelayer.GetDefinition() as CIMFeatureLayer;

                    // Get the label classes - we need the first one.
                    var listLabelClasses = lyrDefn.LabelClasses.ToList();
                    var labelClass = listLabelClasses.FirstOrDefault();

                    // Set the label text symbol.
                    labelClass.TextSymbol.Symbol = textSymbol;

                    // Set the label definition back to the layer.
                    featurelayer.SetDefinition(lyrDefn);

                    // Set the label visibilty.
                    featurelayer.SetLabelVisibility(displayLabels);
                });
            }
            catch
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        public async Task<bool> SwitchLabelsAsync(string layerName, bool displayLabels)
        {
            // Check there is an input layer.
            if (String.IsNullOrEmpty(layerName))
                return false;

            // Get the input feature layer.
            FeatureLayer featurelayer = FindLayer(layerName);

            if (featurelayer == null)
                return false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Set the label visibilty.
                    featurelayer.SetLabelVisibility(displayLabels);
                });
            }
            catch
            {
                // Handle Exception.
                return false;
            }

            return true;
        }

        #endregion Symbology

        #region Export

        public async Task<int> CopyFCToTextFileAsync(string inputLayer, string outputTable, string columns, string orderByColumns,
            bool append = false, bool includeHeader = true)
        {
            // Check there is an input table.
            if (String.IsNullOrEmpty(inputLayer))
                return -1;

            // Check there are columns to output.
            if (String.IsNullOrEmpty(columns))
                return -1;

            // Get the input feature layer.
            FeatureLayer inputFeaturelayer = FindLayer(inputLayer);

            if (inputFeaturelayer == null)
                return -1;

            // Get the list of fields for the input table.
            IReadOnlyList<ArcGIS.Core.Data.Field> inputfields;
            inputfields = await GetFCFieldsAsync(inputLayer);

            // Check a list of fields is returned.
            if (inputfields == null || inputfields.Count == 0)
                return -1;

            // Align the columns with what actually exists in the layer.
            List<string> columnsList = [.. columns.Split(',')];
            bool missingColumns = false;
            columns = "";
            foreach (string column in columnsList)
            {
                string columnName = column.Trim();
                if ((columnName.Substring(0, 1) != "\"") || (!FieldExists(inputfields, columnName)))
                    columns = columns + columnName + ",";
                else
                {
                    missingColumns = true;
                    break;
                }
            }

            // Stop if there are any missing columns.
            if (missingColumns || string.IsNullOrEmpty(columns))
                return -1;
            else
                columns = columns[..^1];

            // Open output file.
            StreamWriter txtFile = new(outputTable, append);

            // Write the header if required.
            if (!append && includeHeader)
                txtFile.WriteLine(columns);

            int intLineCount = 0;
            try
            {
                await QueuedTask.Run(() =>
                {
                    /// Get the feature class for the input feature layer.
                    FeatureClass featureClass = inputFeaturelayer.GetFeatureClass();

                    // Get the feature class defintion.
                    using FeatureClassDefinition featureClassDefinition = featureClass.GetDefinition();

                    // Create a row cursor.
                    RowCursor rowCursor;

                    // Create a new list of sort descriptions.
                    List<ArcGIS.Core.Data.SortDescription> sortDescriptions = [];

                    if (!string.IsNullOrEmpty(orderByColumns))
                    {
                        columnsList = [.. orderByColumns.Split(',')];

                        // Build the list of sort descriptions for each column in the input layer.
                        foreach (string column in columnsList)
                        {
                            string columnName = column.Trim();
                            if ((columnName.Substring(0, 1) != "\"") || (!FieldExists(inputfields, columnName)))
                            {
                                // Get the field from the feature class definition.
                                ArcGIS.Core.Data.Field field = featureClassDefinition.GetFields()
                                  .First(x => x.Name.Equals(columnName));

                                // Create a SortDescription for the field.
                                ArcGIS.Core.Data.SortDescription sortDescription = new(field)
                                {
                                    CaseSensitivity = CaseSensitivity.Insensitive,
                                    SortOrder = ArcGIS.Core.Data.SortOrder.Ascending
                                };

                                // Add the SortDescription to the list.
                                sortDescriptions.Add(sortDescription);
                            }
                        }

                        // Create a TableSortDescription.
                        TableSortDescription tableSortDescription = new(sortDescriptions);

                        // Create a cursor of the sorted features.
                        rowCursor = featureClass.Sort(tableSortDescription);
                    }
                    else
                    {
                        // Create a cursor of the features.
                        rowCursor = featureClass.Search();
                    }

                    // Loop through the feature class/table using the cursor.
                    while (rowCursor.MoveNext())
                    {
                        /// Get the current row.
                        using Row record = rowCursor.Current;

                        string newRow = "";
                        foreach (string column in columnsList)
                        {
                            string columnName = column.Trim();
                            if (columnName.Substring(0, 1) != "\"")
                            {
                                // Get the field value.
                                var columnValue = record[columnName];
                                columnValue ??= "";

                                // Wrap value if quotes if it is a string that contains a comma
                                if ((columnValue is string) && (columnValue.ToString().Contains(',')))
                                    columnValue = "\"" + columnValue.ToString() + "\"";

                                // Format distance to the nearest metre
                                if (columnValue is double && columnName == "Distance")
                                {
                                    double dblValue = double.Parse(columnValue.ToString());
                                    int intValue = Convert.ToInt32(dblValue);
                                    columnValue = intValue;
                                }

                                // Append the column value to the new row.
                                newRow = newRow + columnValue.ToString() + ",";
                            }
                            else
                            {
                                newRow = newRow + columnName + ",";
                            }
                        }

                        // Remove the final comma
                        newRow = newRow[..^1];

                        // Write the new row.
                        txtFile.WriteLine(newRow);
                        intLineCount++;
                    }
                    // Dispose of the objects.
                    featureClass.Dispose();
                    featureClassDefinition.Dispose();
                    rowCursor.Dispose();
                    rowCursor = null;
                });
            }
            catch
            {
                // Handle Exception.
                return 0;
            }
            finally
            {
                // Close the file.
                txtFile.Close();

                // Dispose of the object.
                txtFile.Dispose();
            }

            return intLineCount;
        }

        public async Task<int> CopyTableToTextFileAsync(string inputLayer, string outputTable, string columns, string orderByColumns,
            bool append = false, bool includeHeader = true)
        {
            // Check there is an input table.
            if (String.IsNullOrEmpty(inputLayer))
                return -1;

            // Check there are columns to output.
            if (String.IsNullOrEmpty(columns))
                return -1;

            // Get the input feature layer.
            StandaloneTable inputTable = FindTable(inputLayer);

            if (inputTable == null)
                return -1;

            // Get the list of fields for the input table.
            IReadOnlyList<ArcGIS.Core.Data.Field> inputfields;
            inputfields = await GetTableFieldsAsync(inputLayer);

            // Check a list of fields is returned.
            if (inputfields == null || inputfields.Count == 0)
                return -1;

            // Align the columns with what actually exists in the layer.
            List<string> columnsList = [.. columns.Split(',')];
            bool missingColumns = false;
            columns = "";
            foreach (string column in columnsList)
            {
                string columnName = column.Trim();
                if ((columnName.Substring(0, 1) != "\"") || (!FieldExists(inputfields, columnName)))
                    columns = columns + columnName + ",";
                else
                {
                    missingColumns = true;
                    break;
                }
            }

            // Stop if there are any missing columns;
            if (missingColumns || string.IsNullOrEmpty(columns))
                return -1;
            else
                columns = columns[..^1];

            // Open output file.
            StreamWriter txtFile = new(outputTable, append);

            // Write the header if required.
            if (!append && includeHeader)
                txtFile.WriteLine(columns);

            int intLineCount = 0;
            try
            {
                await QueuedTask.Run(() =>
                {
                    /// Get the underlying table for the input layer.
                    ArcGIS.Core.Data.Table table = inputTable.GetTable();

                    // Get the table defintion.
                    using TableDefinition tableDefinition = table.GetDefinition();

                    // Create a row cursor.
                    RowCursor rowCursor;

                    // Create a new list of sort descriptions.
                    List<ArcGIS.Core.Data.SortDescription> sortDescriptions = [];

                    if (!string.IsNullOrEmpty(orderByColumns))
                    {
                        List<string> orderByColumnsList = [.. orderByColumns.Split(',')];

                        // Build the list of sort descriptions for each column in the input layer.
                        foreach (string column in orderByColumnsList)
                        {
                            string columnName = column.Trim();
                            if ((columnName.Substring(0, 1) != "\"") && (FieldExists(inputfields, columnName)))
                            {
                                // Get the field from the feature class definition.
                                ArcGIS.Core.Data.Field field = tableDefinition.GetFields()
                                  .First(x => x.Name.Equals(columnName));

                                // Create a SortDescription for the field.
                                ArcGIS.Core.Data.SortDescription sortDescription = new(field)
                                {
                                    CaseSensitivity = CaseSensitivity.Insensitive,
                                    SortOrder = ArcGIS.Core.Data.SortOrder.Ascending
                                };

                                // Add the SortDescription to the list.
                                sortDescriptions.Add(sortDescription);
                            }
                        }

                        // Create a TableSortDescription.
                        TableSortDescription tableSortDescription = new(sortDescriptions);

                        // Create a cursor of the sorted features.
                        rowCursor = table.Sort(tableSortDescription);
                    }
                    else
                    {
                        // Create a cursor of the features.
                        rowCursor = table.Search();
                    }

                    // Loop through the feature class/table using the cursor.
                    while (rowCursor.MoveNext())
                    {
                        // Get the current row.
                        using Row record = rowCursor.Current;

                        string newRow = "";
                        foreach (string column in columnsList)
                        {
                            string columnName = column.Trim();
                            if (columnName.Substring(0, 1) != "\"")
                            {
                                // Get the field value.
                                var columnValue = record[columnName];
                                columnValue ??= "";

                                // Wrap value if quotes if it is a string that contains a comma
                                if ((columnValue is string) && (columnValue.ToString().Contains(',')))
                                    columnValue = "\"" + columnValue.ToString() + "\"";

                                // Format distance to the nearest metre
                                if (columnValue is double && columnName == "Distance")
                                {
                                    double dblValue = double.Parse(columnValue.ToString());
                                    int intValue = Convert.ToInt32(dblValue);
                                    columnValue = intValue;
                                }

                                // Append the column value to the new row.
                                newRow = newRow + columnValue.ToString() + ",";
                            }
                            else
                            {
                                newRow = newRow + columnName + ",";
                            }
                        }

                        // Remove the final comma
                        newRow = newRow[..^1];

                        // Write the new row.
                        txtFile.WriteLine(newRow);
                        intLineCount++;
                    }
                    // Dispose of the objects.
                    table.Dispose();
                    tableDefinition.Dispose();
                    rowCursor.Dispose();
                    rowCursor = null;
                });
            }
            catch
            {
                // Handle Exception.
                return 0;
            }
            finally
            {
                // Close the file.
                txtFile.Close();

                // Dispose of the object.
                txtFile.Dispose();
            }

            return intLineCount;
        }

        #endregion Export
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
                // It's an SDE class.
                // Not handled. We know the layer exists.
                return true;
            }
            else // It is a geodatabase class.
            {
                try
                {
                    return await FeatureClassExistsGDBAsync(filePath, fileName);
                }
                catch
                {
                    // GetDefinition throws an exception if the definition doesn't exist.
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
        public static async Task<bool> DeleteGeodatabaseFCAsync(string filePath, string fileName)
        {
            bool success = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Open the file geodatabase. This will open the geodatabase if the folder exists and contains a valid geodatabase.
                    using Geodatabase geodatabase = new(new FileGeodatabaseConnectionPath(new Uri(filePath)));

                    // Create a SchemaBuilder object
                    SchemaBuilder schemaBuilder = new(geodatabase);

                    // Create a FeatureClassDescription object.
                    using FeatureClassDefinition featureClassDefinition = geodatabase.GetDefinition<FeatureClassDefinition>(fileName);

                    // Create a FeatureClassDescription object
                    FeatureClassDescription featureClassDescription = new(featureClassDefinition);

                    // Add the deletion for the feature class to the list of DDL tasks
                    schemaBuilder.Delete(featureClassDescription);

                    // Execute the DDL
                    success = schemaBuilder.Build();
                });
            }
            catch (GeodatabaseNotFoundOrOpenedException)
            {
                // Handle Exception.
                return false;
            }
            catch (GeodatabaseTableException)
            {
                // Handle Exception.
                return false;
            }

            return success;
        }

        /// <summary>
        /// Delete a feature class from a geodatabase.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="featureClassName"></param>
        public static async Task<bool> DeleteGeodatabaseFCAsync(Geodatabase geodatabase, string featureClassName)
        {
            bool success = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Create a SchemaBuilder object
                    SchemaBuilder schemaBuilder = new(geodatabase);

                    // Create a FeatureClassDescription object.
                    using FeatureClassDefinition featureClassDefinition = geodatabase.GetDefinition<FeatureClassDefinition>(featureClassName);

                    // Create a FeatureClassDescription object
                    FeatureClassDescription featureClassDescription = new(featureClassDefinition);

                    // Add the deletion for the feature class to the list of DDL tasks
                    schemaBuilder.Delete(featureClassDescription);

                    // Execute the DDL
                    success = schemaBuilder.Build();
                });
            }
            catch
            {
                // Handle exception.
                return false;
            }

            return success;
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
            if (String.IsNullOrEmpty(fileName))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(fileName);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; //| GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.Delete", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.Delete", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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

        public static async Task<bool> AddFieldAsync(string inTable, string fieldName, string fieldType = "TEXT",
            long fieldPrecision = -1, long fieldScale = -1, long fieldLength = -1, string fieldAlias = null,
            bool fieldIsNullable = true, bool fieldIsRequred = false, string fieldDomain = null)
        {
            if (String.IsNullOrEmpty(inTable))
                return false;

            if (String.IsNullOrEmpty(fieldName))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, fieldName, fieldType,
                fieldPrecision > 0 ? fieldPrecision : null, fieldScale > 0 ? fieldScale : null, fieldLength > 0 ? fieldLength : null,
                fieldAlias ?? null, fieldIsNullable ? "NULLABLE" : "NON_NULLABLE",
                fieldIsRequred ? "REQUIRED" : "NON_REQUIRED", fieldDomain);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; //| GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.AddField", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.AddField", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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

        public static async Task<bool> RenameFieldAsync(string inTable, string fieldName, string newFieldName)
        {
            if (String.IsNullOrEmpty(inTable))
                return false;

            if (String.IsNullOrEmpty(fieldName))
                return false;

            if (String.IsNullOrEmpty(newFieldName))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, fieldName, newFieldName);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; //| GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.AlterField", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.AlterField", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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

        public static async Task<bool> CalculateFieldAsync(string inTable, string fieldName, string fieldCalc)
        {
            if (String.IsNullOrEmpty(inTable))
                return false;

            if (String.IsNullOrEmpty(fieldName))
                return false;

            if (String.IsNullOrEmpty(fieldCalc))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, fieldName, fieldCalc);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; //| GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.CalculateField", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.CalculateField", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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

        public static async Task<bool> CalculateGeometryAsync(string inTable, string geometryProperty, string lineUnit = "", string areaUnit = "")
        {
            if (String.IsNullOrEmpty(inTable))
                return false;

            if (String.IsNullOrEmpty(geometryProperty))
                return false;

            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, geometryProperty, lineUnit, areaUnit);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; //| GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.CalculateGeometryAttributes", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.CalculateGeometryAttributes", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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

        /// <summary>
        /// Select features in layerName by location.
        /// </summary>
        /// <param name="layerName"></param>
        /// <param name="whereClause"></param>
        /// <returns></returns>
        public static async Task<bool> SelectLayerByLocationAsync(string targetLayer, string searchLayer,
            string overlapType = "INTERSECT", string searchDistance = "", string selectionType = "NEW_SELECTION")
        {
            if (String.IsNullOrEmpty(targetLayer))
                return false;

            if (String.IsNullOrEmpty(searchLayer))
                return false;

            // Make a value array of strings to be passed to the tool.
            IReadOnlyList<string> parameters = Geoprocessing.MakeValueArray(targetLayer, overlapType, searchLayer, searchDistance, selectionType);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;

            //Geoprocessing.OpenToolDialog("management.SelectLayerByLocation", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.SelectLayerByLocation", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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

        public static async Task<bool> BufferFeaturesAsync(string inFeatureClass, string outFeatureClass, string bufferDistance,
            string lineSide = "FULL", string lineEndType = "ROUND", string dissolveOption = "NONE", string dissolveFields = "", string method = "PLANAR", bool addToMap = false)
        {
            if (String.IsNullOrEmpty(inFeatureClass))
                return false;

            if (String.IsNullOrEmpty(outFeatureClass))
                return false;

            if (String.IsNullOrEmpty(bufferDistance))
                return false;

            // Make a value array of strings to be passed to the tool.
            //List<string> parameters = [.. Geoprocessing.MakeValueArray(inFeatureClass, outFeatureClass, bufferDistance, lineSide, lineEndType, method, dissolveOption)];
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inFeatureClass, outFeatureClass, bufferDistance, lineSide, lineEndType, dissolveOption)];
            if (!string.IsNullOrEmpty(dissolveFields))
                parameters.Add(dissolveFields);
            parameters.Add(method);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.Buffer", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Buffer", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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

        public static async Task<bool> ClipFeaturesAsync(string inFeatureClass, string clipFeatureClass, string outFeatureClass, bool addToMap = false)
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
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.Clip", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Clip", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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

        public static async Task<bool> IntersectFeaturesAsync(string inFeatures, string outFeatureClass, string joinAttributes = "ALL", string outputType = "INPUT", bool addToMap = false)
        {
            if (String.IsNullOrEmpty(inFeatures))
                return false;

            if (String.IsNullOrEmpty(outFeatureClass))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inFeatures, outFeatureClass, joinAttributes, outputType)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.Intersect", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Intersect", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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

        public static async Task<bool> SpatialJoinAsync(string targetFeatures, string joinFeatures, string outFeatureClass, string joinOperation = "JOIN_ONE_TO_ONE",
            string joinType = "KEEP_ALL", string fieldMapping = "", string matchOption = "INTERSECT", string searchRadius = "0", string distanceField = "",
            string matchFields = "", bool addToMap = false)
        {
            if (String.IsNullOrEmpty(targetFeatures))
                return false;

            if (String.IsNullOrEmpty(joinFeatures))
                return false;

            if (String.IsNullOrEmpty(outFeatureClass))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(targetFeatures, joinFeatures, outFeatureClass, joinOperation, joinType, fieldMapping,
                matchOption, searchRadius, distanceField, matchFields)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.SpatialJoin", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.SpatialJoin", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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

        public static async Task<bool> CalculateSummaryStatisticsAsync(string inTable, string outTable, string statisticsFields,
            string caseFields = "", string concatenationSeparator = "", bool addToMap = false)
        {
            if (String.IsNullOrEmpty(inTable))
                return false;

            if (String.IsNullOrEmpty(outTable))
                return false;

            if (String.IsNullOrEmpty(statisticsFields))
                return false;

            // Make a value array of strings to be passed to the tool.
            List<string> parameters = [.. Geoprocessing.MakeValueArray(inTable, outTable, statisticsFields, caseFields, concatenationSeparator)];

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("analysis.Statistics", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("analysis.Statistics", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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

        #endregion Feature Class

        #region Geodatabase

        public static Geodatabase CreateFileGeodatabase(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath))
                return null;

            Geodatabase geodatabase;

            try
            {
                // Create a FileGeodatabaseConnectionPath with the name of the file geodatabase you wish to create
                FileGeodatabaseConnectionPath fileGeodatabaseConnectionPath = new(new Uri(fullPath));

                // Create and use the file geodatabase
                geodatabase = SchemaBuilder.CreateGeodatabase(fileGeodatabaseConnectionPath);
            }
            catch
            {
                // Handle Exception.
                return null;
            }

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
            catch (GeodatabaseTableException)
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
            catch (GeodatabaseTableException)
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
                    // GetDefinition throws an exception if the definition doesn't exist.
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
                // It's an SDE class.
                // Not handled. We know the layer exists.
                return true;
            }
            else // It is a geodatabase class.
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
        public static async Task<bool> DeleteGeodatabaseTableAsync(string filePath, string fileName)
        {
            if (string.IsNullOrEmpty(filePath))
                return false;

            if (string.IsNullOrEmpty(fileName))
                return false;

            bool success = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    // Open the file geodatabase. This will open the geodatabase if the folder exists and contains a valid geodatabase.
                    using Geodatabase geodatabase = new(new FileGeodatabaseConnectionPath(new Uri(filePath)));

                    // Create a SchemaBuilder object
                    SchemaBuilder schemaBuilder = new(geodatabase);

                    // Create a FeatureClassDescription object.
                    using TableDefinition tableDefinition = geodatabase.GetDefinition<TableDefinition>(fileName);

                    // Create a FeatureClassDescription object
                    TableDescription tableDescription = new(tableDefinition);

                    // Add the deletion for the feature class to the list of DDL tasks
                    schemaBuilder.Delete(tableDescription);

                    // Execute the DDL
                    success = schemaBuilder.Build();
                });
            }
            catch
            {
                return false;
            }

            return success;
        }

        /// <summary>
        /// Delete a table from a geodatabase.
        /// </summary>
        /// <param name="geodatabase"></param>
        /// <param name="tableName"></param>
        public static async Task<bool> DeleteGeodatabaseTableAsync(Geodatabase geodatabase, string tableName)
        {
            if (geodatabase == null)
                return false;

            if (string.IsNullOrEmpty(tableName))
                return false;

            bool success = false;

            try
            {
                await QueuedTask.Run(() =>
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
                    success = schemaBuilder.Build();
                });
            }
            catch
            {
                return false;
            }

            return success;
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
            BrowseProjectFilter bf = fileType switch
            {
                "Geodatabase FC" => BrowseProjectFilter.GetFilter("esri_browseDialogFilters_geodatabaseItems_featureClasses"),
                "Geodatabase Table" => BrowseProjectFilter.GetFilter("esri_browseDialogFilters_geodatabaseItems_tables"),
                "Shapefile"=> BrowseProjectFilter.GetFilter("esri_browseDialogFilters_shapefiles"),
                "CSV file (comma delimited)"=> BrowseProjectFilter.GetFilter("esri_browseDialogFilters_textFiles_csv"),
                "Text file (tab delimited)" => BrowseProjectFilter.GetFilter("esri_browseDialogFilters_textFiles_txt"),
                _ => BrowseProjectFilter.GetFilter("esri_browseDialogFilters_all"),
            };

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
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("management.CopyFeatures", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.CopyFeatures", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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
        /// <param name="inputWorkspace"></param>
        /// <param name="inputDatasetName"></param>
        /// <param name="outputFeatureClass"></param>
        /// <param name="Messages"></param>
        /// <returns></returns>
        public static async Task<bool> CopyFeaturesAsync(string inputWorkspace, string inputDatasetName, string outputFeatureClass)
        {
            string inFeatureClass = inputWorkspace + @"\" + inputDatasetName;

            return await CopyFeaturesAsync(inFeatureClass, outputFeatureClass);
        }

        /// <summary>
        /// Copy the input dataset to the output dataset.
        /// </summary>
        /// <param name="inputWorkspace"></param>
        /// <param name="inputDatasetName"></param>
        /// <param name="outputWorkspace"></param>
        /// <param name="outputDatasetName"></param>
        /// <param name="Messages"></param>
        /// <returns></returns>
        public static async Task<bool> CopyFeaturesAsync(string inputWorkspace, string inputDatasetName, string outputWorkspace, string outputDatasetName)
        {
            string inFeatureClass = inputWorkspace + @"\" + inputDatasetName;
            string outFeatureClass = outputWorkspace + @"\" + outputDatasetName;

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
        public static async Task<bool> ExportFeaturesAsync(string inTable, string outTable, bool addToMap = false)
        {
            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, outTable);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("conversion.ExportTable", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("conversion.ExportTable", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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
        public static async Task<bool> CopyTableAsync(string inTable, string outTable, bool addToMap = false)
        {
            // Make a value array of strings to be passed to the tool.
            var parameters = Geoprocessing.MakeValueArray(inTable, outTable);

            // Make a value array of the environments to be passed to the tool.
            var environments = Geoprocessing.MakeEnvironmentArray(overwriteoutput: true);

            // Set the geprocessing flags.
            GPExecuteToolFlags executeFlags = GPExecuteToolFlags.GPThread; // | GPExecuteToolFlags.RefreshProjectItems;
            if (addToMap)
                executeFlags |= GPExecuteToolFlags.AddOutputsToMap;

            //Geoprocessing.OpenToolDialog("management.Copy", parameters);  // Useful for debugging.

            // Execute the tool.
            try
            {
                IGPResult gp_result = await Geoprocessing.ExecuteToolAsync("management.Copy", parameters, environments, null, null, executeFlags);

                if (gp_result.IsFailed)
                {
                    Geoprocessing.ShowMessageBox(gp_result.Messages, "GP Messages", GPMessageBoxStyle.Error);

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