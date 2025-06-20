﻿// The DataTools are a suite of ArcGIS Pro addins used to extract, sync
// and manage biodiversity information from ArcGIS Pro and SQL Server
// based on pre-defined or user specified criteria.
//
// Copyright © 2024-25 Andy Foy Consulting.
//
// This file is part of DataTools suite of programs.
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

using ArcGIS.Core.Data;
using ArcGIS.Core.Data.UtilityNetwork.Trace;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Framework.Controls;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using DataTools;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

namespace DataSearches.UI
{
    internal class PaneHeader2ViewModel : PanelViewModelBase, INotifyPropertyChanged
    {
        #region Enums

        /// <summary>
        /// An enumeration of the different options for whether to
        /// add output layers to the map.
        /// </summary>
        public enum AddSelectedLayersOptions
        {
            No,
            WithoutLabels,
            WithLabels
        };

        /// <summary>
        /// An enumeration of the different options for whether to
        /// overwrite the label for each output layer.
        /// </summary>
        public enum OverwriteLabelOptions
        {
            No,
            ResetByLayer,
            ResetByGroup,
            DoNotReset
        };

        /// <summary>
        /// An enumeration of the different options for whether to
        /// overwrite the label for each output layer.
        /// </summary>
        public enum CombinedSitesTableOptions
        {
            None,
            Append,
            Overwrite
        };

        #endregion Enums

        #region Fields

        private readonly DockpaneMainViewModel _dockPane;

        private string _userID;
        private string _tempGDBName;

        private bool _searchErrors;

        private bool _updateTable;

        private string _repChar;

        private string _databasePath;
        private string _databaseTable;
        private string _databaseRefColumn;
        private string _databaseSiteColumn;
        private string _databaseOrgColumn;

        private bool _requireSiteName;
        private bool _requireOrganisation;

        private bool? _defaultKeepSelectedLayers;
        private int _defaultAddSelectedLayers;
        private int _defaultOverwriteLabels;
        private int _defaultCombinedSitesTable;

        private readonly List<string> _bufferUnitOptionsDisplay = ["Centimetres", "Metres", "Kilometres", "Feet", "Yards", "Miles"];
        private readonly List<string> _bufferUnitOptionsProcess = ["Centimeters", "Meters", "Kilometers", "Feet", "Yards", "Miles"];
        private readonly List<string> _bufferUnitOptionsShort = ["cm", "m", "km", "ft", "yds", "mi"];

        private List<MapLayer> _mapLayers;

        private List<MapLayer> _openMapLayersList;
        private List<string> _closedMapLayersList;

        private string _saveRootDir;
        private string _saveFolder;
        private string _extractFolder;
        private string _gisFolder;
        private string _outputPath;
        private string _gisPath;
        private string _layerPath;

        private string _logFileName;
        private string _combinedSitesTableName;
        private string _combinedSitesColumnList;
        private string _combinedSitesTableFormat;
        private string _combinedSitesOutputFile;

        private string _bufferPrefix;
        private string _bufferLayerName;
        private FeatureLayer _bufferLayer;
        private string _bufferLayerPath;
        private string _bufferOutputFile;
        private string _bufferLayerFile;
        private string _bufferFields;
        private bool _keepBuffer;

        private string _searchLayerName;
        private string _searchOutputFile;
        private string _groupLayerName;

        private List<string> _keepSearchFeatureExtensions;
        private string _searchSymbologyBase;

        private string _searchLayerBase;
        private List<string> _searchLayerExtensions;
        private string _searchLayerExtension;
        private string _inputLayerName;

        private string _searchColumn;
        private string _siteColumn;
        private string _orgColumn;
        private string _radiusColumn;

        private List<string> _mapGroupNames = [];
        private List<int> _mapGroupLabels = [];
        private int _maxLabel = 1;

        private string _logFile;

        private string _tempMasterLayerName;
        private string _tempMasterOutputFile;
        private string _tempFCLayerName;
        private string _tempFCOutputFile;
        private string _tempTableLayerName;
        private string _tempTableOutputFile;

        private string _tempFCPointsLayerName;
        private string _tempFCPointsOutputFile;
        private string _tempSearchPointsLayerName;
        private string _tempSearchPointsOutputFile;

        private const string _displayName = "DataSearches";

        private readonly DataSearchesConfig _toolConfig;
        private MapFunctions _mapFunctions;

        #endregion Fields

        #region ViewModelBase Members

        public override string DisplayName
        {
            get { return _displayName; }
        }

        #endregion ViewModelBase Members

        #region Creator

        /// <summary>
        /// Set the global variables.
        /// </summary>
        /// <param name="xmlFilesList"></param>
        /// <param name="defaultXMLFile"></param>
        public PaneHeader2ViewModel(DockpaneMainViewModel dockPane, DataSearchesConfig toolConfig)
        {
            _dockPane = dockPane;

            // Return if no config object passed.
            if (toolConfig == null) return;

            // Set the config object.
            _toolConfig = toolConfig;

            InitializeComponent();
        }

        /// <summary>
        /// Initialise the search pane.
        /// </summary>
        private void InitializeComponent()
        {
            // Get the relevant config file settings.
            _repChar = _toolConfig.RepChar;

            _databasePath = _toolConfig.DatabasePath;
            _databaseTable = _toolConfig.DatabaseTable;
            _databaseRefColumn = _toolConfig.DatabaseRefColumn;
            _databaseSiteColumn = _toolConfig.DatabaseSiteColumn;
            _databaseOrgColumn = _toolConfig.DatabaseOrgColumn;

            _requireSiteName = _toolConfig.RequireSiteName;
            _requireOrganisation = _toolConfig.RequireOrganisation;

            _defaultKeepSelectedLayers = _toolConfig.DefaultKeepSelectedLayers;
            _defaultAddSelectedLayers = _toolConfig.DefaultAddSelectedLayers;
            _defaultOverwriteLabels = _toolConfig.DefaultOverwriteLabels;
            _defaultCombinedSitesTable = _toolConfig.DefaultCombinedSitesTable;

            _updateTable = _toolConfig.UpdateTable;

            _saveRootDir = _toolConfig.SaveRootDir;
            _saveFolder = _toolConfig.SaveFolder;
            _extractFolder = _toolConfig.ExtractFolder;
            _gisFolder = _toolConfig.GISFolder;
            _logFileName = _toolConfig.LogFileName;
            _layerPath = _toolConfig.LayerFolder;
            _combinedSitesTableName = _toolConfig.CombinedSitesTableName;
            _combinedSitesColumnList = _toolConfig.CombinedSitesTableColumns;
            _combinedSitesTableFormat = _toolConfig.CombinedSitesTableFormat;

            _bufferPrefix = _toolConfig.BufferPrefix;
            _bufferLayerFile = _toolConfig.BufferLayerFile;
            _searchLayerName = _toolConfig.SearchOutputName;
            _groupLayerName = _toolConfig.GroupLayerName;

            _searchLayerBase = _toolConfig.SearchLayer;
            _searchLayerExtensions = _toolConfig.SearchLayerExtensions;

            _searchSymbologyBase = _toolConfig.SearchSymbologyBase;
            _keepSearchFeatureExtensions = _toolConfig.KeepSearchFeatureExtensions;

            _searchColumn = _toolConfig.SearchColumn;
            _siteColumn = _toolConfig.SiteColumn;
            _radiusColumn = _toolConfig.RadiusColumn;
            _orgColumn = _toolConfig.OrgColumn;

            _bufferFields = _toolConfig.AggregateColumns;
            _keepBuffer = _toolConfig.KeepBufferArea;
        }

        #endregion Creator

        #region Controls Enabled

        /// <summary>
        /// Is the list of layers enabled?
        /// </summary>
        public bool LayersListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_mapLayersList != null));
            }
        }

        /// <summary>
        /// Is the list of buffer units options enabled?
        /// </summary>
        public bool BufferUnitsListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_bufferUnitsList != null));
            }
        }

        /// <summary>
        /// Is the option to keep the selected layers enabled?
        /// </summary>
        public bool KeepSelectedLayersEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_defaultKeepSelectedLayers != null));
            }
        }

        /// <summary>
        /// Is the list of add to map options enabled?
        /// </summary>
        public bool AddToMapListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_addToMapList != null));
            }
        }

        /// <summary>
        /// Is the list of overwrite options enabled?
        /// </summary>
        public bool OverwriteLabelsListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_overwriteLabelsList != null));
            }
        }

        /// <summary>
        /// Is the list of combined sites options enabled?
        /// </summary>
        public bool CombinedSitesListEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_combinedSitesList != null));
            }
        }

        /// <summary>
        /// Can the run button be pressed?
        /// </summary>
        public bool RunButtonEnabled
        {
            get
            {
                return ((_dockPane.ProcessStatus == null)
                    && (_mapLayersList != null)
                    && (_mapLayersList.Where(p => p.IsSelected).Any())
                    && (!string.IsNullOrEmpty(_searchRefText))
                    && (!_requireSiteName || !string.IsNullOrEmpty(_siteNameText))
                    && (!_requireOrganisation || !string.IsNullOrEmpty(_organisationText))
                    && (!string.IsNullOrEmpty(_bufferSizeText))
                    && (_selectedBufferUnitsIndex >= 0)
                    && (_defaultAddSelectedLayers <= 0 || _selectedAddToMap != null)
                    && (_selectedOverwriteLabels != null)
                    && (_defaultCombinedSitesTable <= 0 || _selectedCombinedSites != null));
            }
        }

        #endregion Controls Enabled

        #region Controls Visibility

        /// <summary>
        /// Is the site name text box visible.
        /// </summary>
        public Visibility SiteNameTextVisibility
        {
            get
            {
                if (!_requireSiteName)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the organisation text box visible.
        /// </summary>
        public Visibility OrganisationTextVisibility
        {
            get
            {
                if (!_requireOrganisation)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the LayersList expand button visible.
        /// </summary>
        public Visibility MapLayersListExpandButtonVisibility
        {
            get
            {
                if ((_mapLayersList == null) || (_mapLayersList.Count < 20))
                    return Visibility.Hidden;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the option to keep output layers visible?
        /// </summary>
        public Visibility KeepLayersVisibility
        {
            get
            {
                if (_defaultKeepSelectedLayers == null)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the list of add to map options visible?
        /// </summary>
        public Visibility AddToMapListVisibility
        {
            get
            {
                if (KeepSelectedLayers == false)
                    return Visibility.Collapsed;
                else
                {
                    if (_defaultAddSelectedLayers == -1)
                        return Visibility.Collapsed;
                    else
                        return Visibility.Visible;
                }
            }
        }

        /// <summary>
        /// Is the list of overwrite options visible?
        /// </summary>
        public Visibility OverwriteLabelsListVisibility
        {
            get
            {
                if (KeepSelectedLayers == false)
                    return Visibility.Collapsed;
                else
                {
                    if (_defaultOverwriteLabels == -1)
                        return Visibility.Collapsed;
                    else
                    {
                        try
                        {
                            int index = AddToMapList.FindIndex(a => a == SelectedAddToMap);
                            if (index > 0)
                                return Visibility.Visible;
                            else
                                return Visibility.Collapsed;
                        }
                        catch
                        {
                            return Visibility.Collapsed;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Is the list of combined sites options visible?
        /// </summary>
        public Visibility CombinedSitesListVisibility
        {
            get
            {
                if (_defaultCombinedSitesTable == -1)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        #endregion Controls Visibility

        #region Message

        private string _message;

        /// <summary>
        /// The message to display on the form.
        /// </summary>
        public string Message
        {
            get
            {
                return _message;
            }
            set
            {
                _message = value;
                OnPropertyChanged(nameof(HasMessage));
                OnPropertyChanged(nameof(Message));
            }
        }

        private MessageType _messageLevel;

        /// <summary>
        /// The type of message; Error, Warning, Confirmation, Information
        /// </summary>
        public MessageType MessageLevel
        {
            get
            {
                return _messageLevel;
            }
            set
            {
                _messageLevel = value;
                OnPropertyChanged(nameof(MessageLevel));
            }
        }

        /// <summary>
        /// Is there a message to display?
        /// </summary>
        public Visibility HasMessage
        {
            get
            {
                if (_dockPane.ProcessStatus != null
                || string.IsNullOrEmpty(_message))
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Show the message with the required icon (message type).
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="messageLevel"></param>
        public void ShowMessage(string msg, MessageType messageLevel)
        {
            MessageLevel = messageLevel;
            Message = msg;
        }

        /// <summary>
        /// Clear the form messages.
        /// </summary>
        public void ClearMessage()
        {
            Message = "";
        }

        #endregion Message

        #region Run Command

        /// <summary>
        /// Validates and executes the search.
        /// </summary>
        public async void ProcessSearchAsync()
        {
            // Indicate the search is initialising.
            _dockPane.SearchStatus = DockpaneMainViewModel.SearchStatuses.Initialising;

            // Validate the parameters.
            if (!ValidateParameters())
                return;

            // Create the search query.
            string searchRef = SearchRefText;
            string searchClause = _searchColumn + " = '" + searchRef + "'";

            // Find the features matching the search reference.
            if (!await FindSearchFeaturesAsync(searchClause))
            {
                ShowMessage("Search ref not found in search layer(s).", MessageType.Warning);
                return;
            }

            // Clear any messages.
            ClearMessage();

            // Update the fields and buttons in the form.
            UpdateFormControls();
            _dockPane.RefreshPanel1Buttons();

            // Reset the map functions object.
            _mapFunctions = null;

            // Process the search.
            bool success = await RunSearchAsync();

            // Indicate that the search process has completed (successfully or not).
            string message;
            string image;
            if (success && !_searchErrors)
            {
                message = "Search '{0}' complete!";
                image = "Success";
            }
            else if (_searchErrors)
            {
                message = "Search '{0}' ended with errors!";
                image = "Error";
            }
            else if (_dockPane.SearchStatus == DockpaneMainViewModel.SearchStatuses.Cancelled)
            {
                message = "Search '{0}' cancelled!";
                image = "Warning";
            }
            else
            {
                message = "Search '{0}' ended unexpectedly!";
                image = "Error";
            }

            // Resume the active map redrawing.
            _mapFunctions.PauseDrawing(false);

            // Indicate that the search process has stopped.
            StopSearch(string.Format(message, searchRef), image);

            // Update the fields and buttons in the form.
            UpdateFormControls();
            _dockPane.RefreshPanel1Buttons();

            // Flushes and closes all trace listeners.
            CloseTraceLogging();
        }

        /// <summary>
        /// Validate the form parameters.
        /// </summary>
        /// <returns></returns>
        private bool ValidateParameters()
        {
            // Site ref is mandatory.
            if (string.IsNullOrEmpty(SearchRefText))
            {
                ShowMessage("Please enter a search reference.", MessageType.Warning);
                return false;
            }

            // Site name is not always required.
            if (_requireSiteName && string.IsNullOrEmpty(SiteNameText))
            {
                ShowMessage("Please enter a site name.", MessageType.Warning);
                return false;
            }

            // Organisation is not always required.
            if (_requireOrganisation && string.IsNullOrEmpty(OrganisationText))
            {
                ShowMessage("Please enter an organisation.", MessageType.Warning);
                return false;
            }

            // At least one layer must be selected,
            if (!MapLayersList.Where(p => p.IsSelected).Any())
            {
                ShowMessage("Please select at least one layer to search.", MessageType.Warning);
                return false;
            }

            // A buffer size must be entered.
            if (string.IsNullOrEmpty(BufferSizeText))
            {
                ShowMessage("Please enter a buffer size.", MessageType.Warning);
                return false;
            }

            // The buffer size must be numeric and positive.
            bool bufferNumeric = double.TryParse(BufferSizeText, out double bufferNumber);
            if (!bufferNumeric || bufferNumber < 0) // User either entered text or a negative number
            {
                ShowMessage("Please enter a positive number for the buffer size.", MessageType.Warning);
                return false;
            }

            // A buffer unit must be selected.
            if (SelectedBufferUnitsIndex < 0)
            {
                ShowMessage("Please select a buffer unit.", MessageType.Warning);
                return false;
            }

            // An add to map option must be selected (if visible).
            if (_defaultAddSelectedLayers != -1)
            {
                if (string.IsNullOrEmpty(SelectedAddToMap))
                {
                    ShowMessage("Please select whether layers should be added to the map.", MessageType.Warning);
                    return false;
                }
            }

            // An overwrite labels option must be selected (if visible).
            if (!SelectedAddToMap.Equals("no", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrEmpty(SelectedOverwriteLabels))
                {
                    ShowMessage("Please select whether to overwrite labels for map layers.", MessageType.Warning);
                    return false;
                }
            }

            // A combined sites table option must be selected (if visible).
            if (_defaultCombinedSitesTable != -1)
            {
                if (string.IsNullOrEmpty(SelectedCombinedSites))
                {
                    ShowMessage("Please select whether the combined sites table should be created.", MessageType.Warning);
                    return false;
                }
            }

            ClearMessage();
            return true;
        }

        #endregion Run Command

        #region Properties

        private string _searchRefText;

        /// <summary>
        /// Get/Set the search reference.
        /// </summary>
        public string SearchRefText
        {
            get
            {
                return _searchRefText;
            }
            set
            {
                _searchRefText = value;

                // Create the search query.
                string searchRef = SearchRefText;
                string searchClause = _searchColumn + " = '" + searchRef + "'";

                if (!string.IsNullOrEmpty(_searchRefText) && _searchRefText.Length > 2)
                {
                    // Find the features matching the search reference.
                    QueuedTask.Run(() =>
                    {
                        if (!FindSearchFeaturesAsync(searchClause).Result)
                            ShowMessage("Search ref not found in search layer(s).", MessageType.Warning);
                        else
                            ClearMessage();
                    });
                }

                //TODO:
                //// If we have a database path and the site name or organisation
                //// are required then try and look them up.
                //if (!string.IsNullOrEmpty(_databasePath) &&  ((_requireSiteName) || (_requireOrganisation)))
                //{
                //    if (!string.IsNullOrEmpty(_searchRefText) && _searchRefText.Length > 2)
                //    {
                //        string siteName = null;
                //        string organisation = null;
                //        if (LookupSearchRef(_searchRefText, ref siteName, ref organisation))
                //        {
                //            //TODO: Set site name and organisation in form.
                //        }
                //        else
                //            ShowMessage("Search ref not found in database", MessageType.Warning);
                //    }
                //}

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private string _siteNameText;

        /// <summary>
        /// Get/Set the search site name.
        /// </summary>
        public string SiteNameText
        {
            get
            {
                return _siteNameText;
            }
            set
            {
                _siteNameText = value;

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        /// <summary>
        /// The tool tip for the site name textbox.
        /// </summary>
        public static string SiteNameTooltip
        {
            get
            {
                //if (_dockPane.LayersListLoading ???
                //|| _selectedTable == null)
                //    return null;
                //else
                return "Enter site name for search";
            }
        }

        private string _organisationText;

        /// <summary>
        /// Get/Set the search organisation.
        /// </summary>
        public string OrganisationText
        {
            get
            {
                return _organisationText;
            }
            set
            {
                _organisationText = value;

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        /// <summary>
        /// The tool tip for the organisation textbox.
        /// </summary>
        public static string OrganisationTooltip
        {
            get
            {
                //if (_dockPane.LayersListLoading ???
                //|| _selectedTable == null)
                //    return null;
                //else
                return "Enter organisation for search";
            }
        }

        private ObservableCollection<MapLayer> _mapLayersList;

        /// <summary>
        /// Get the list of loaded GIS layers.
        /// </summary>
        public ObservableCollection<MapLayer> MapLayersList
        {
            get
            {
                return _mapLayersList;
            }
            set
            {
                _mapLayersList = value;
                OnPropertyChanged(nameof(MapLayersListExpandButtonVisibility));
            }
        }

        private double? _mapLayersListHeight = null;

        public double? MapLayersListHeight
        {
            get
            {
                if (_mapLayersList == null || _mapLayersList.Count == 0)
                    return 382;
                else
                    return _mapLayersListHeight;
            }
        }

        public string MapLayersListExpandButtonContent
        {
            get
            {
                if (_mapLayersListHeight == null)
                    return "-";
                else
                    return "+";
            }
        }

        /// <summary>
        /// Triggered when the selection in the list of layers changes.
        /// </summary>
        public int MapLayersList_SelectedIndex
        {
            set
            {
                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private List<MapLayer> _selectedLayers;

        /// <summary>
        /// Get/Set the selected loaded layers.
        /// </summary>
        public List<MapLayer> SelectedLayers
        {
            get
            {
                return _selectedLayers;
            }
            set
            {
                _selectedLayers = value;
            }
        }

        private string _bufferSizeText;

        /// <summary>
        /// Get/Set the buffer size text.
        /// </summary>
        public string BufferSizeText
        {
            get
            {
                return _bufferSizeText;
            }
            set
            {
                _bufferSizeText = value;

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private List<string> _bufferUnitsList;

        /// <summary>
        /// Get/Set the buffer size units.
        /// </summary>
        public List<string> BufferUnitsList
        {
            get
            {
                return _bufferUnitsList;
            }
            set
            {
                _bufferUnitsList = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(BufferUnitsListEnabled));
            }
        }

        private int _selectedBufferUnitsIndex;

        /// <summary>
        /// Get/Set the selected buffer size units index.
        /// </summary>
        public int SelectedBufferUnitsIndex
        {
            get
            {
                return _selectedBufferUnitsIndex;
            }
            set
            {
                _selectedBufferUnitsIndex = value;

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private bool _keepSelectedLayers;

        /// <summary>
        /// Get/Set the option for whether to keep the selected output
        /// layers.
        /// </summary>
        public bool KeepSelectedLayers
        {
            get
            {
                return _keepSelectedLayers;
            }
            set
            {
                _keepSelectedLayers = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(KeepSelectedLayersEnabled));
                OnPropertyChanged(nameof(AddToMapListVisibility));
                OnPropertyChanged(nameof(OverwriteLabelsListVisibility));
            }
        }

        private List<string> _addToMapList = ["No", "Yes - Without labels", "Yes - With labels"];

        /// <summary>
        /// Get/Set the options for whether the output layers will be added
        /// to the map.
        /// </summary>
        public List<string> AddToMapList
        {
            get
            {
                return _addToMapList;
            }
            set
            {
                _addToMapList = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(AddToMapListEnabled));
            }
        }

        private string _selectedAddToMap;

        /// <summary>
        /// Get/Set whether the output layers will be added to the map.
        /// </summary>
        public string SelectedAddToMap
        {
            get
            {
                return _selectedAddToMap;
            }
            set
            {
                _selectedAddToMap = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(OverwriteLabelsListVisibility));

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private List<string> _overwriteLabelsList = ["No", "Yes - Reset each layer", "Yes - Reset each group", "Yes - Do not reset"];

        /// <summary>
        /// Get/Set the search site name.
        /// </summary>
        public List<string> OverwriteLabelsList
        {
            get
            {
                return _overwriteLabelsList;
            }
            set
            {
                _overwriteLabelsList = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(OverwriteLabelsListEnabled));
            }
        }

        private string _selectedOverwriteLabels;

        /// <summary>
        /// Get/Set the search site name.
        /// </summary>
        public string SelectedOverwriteLabels
        {
            get
            {
                return _selectedOverwriteLabels;
            }
            set
            {
                _selectedOverwriteLabels = value;

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private List<string> _combinedSitesList = ["None", "Append to existing table", "Overwrite existing table"];

        /// <summary>
        /// Get/Set the search site name.
        /// </summary>
        public List<string> CombinedSitesList
        {
            get
            {
                return _combinedSitesList;
            }
            set
            {
                _combinedSitesList = value;

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(CombinedSitesListEnabled));
            }
        }

        private string _selectedCombinedSites;

        /// <summary>
        /// Get/Set the search site name.
        /// </summary>
        public string SelectedCombinedSites
        {
            get
            {
                return _selectedCombinedSites;
            }
            set
            {
                _selectedCombinedSites = value;

                // Check if the run button is now enabled/disabled.
                _dockPane.CheckRunButton();
            }
        }

        private bool _clearLogFile;

        /// <summary>
        /// Is the log file to be cleared before running the search?
        /// </summary>
        public bool ClearLogFile
        {
            get
            {
                return _clearLogFile;
            }
            set
            {
                _clearLogFile = value;
            }
        }

        private bool _openLogFile;

        /// <summary>
        /// Is the log file to be opened after running the search?
        /// </summary>
        public bool OpenLogFile
        {
            get
            {
                return _openLogFile;
            }
            set
            {
                _openLogFile = value;
            }
        }

        private bool _pauseMap;

        /// <summary>
        /// Whether the map processing should be paused during processing?
        /// </summary>
        public bool PauseMap
        {
            get
            {
                return _pauseMap;
            }
            set
            {
                _pauseMap = value;
            }
        }

        #endregion Properties

        #region Debug Logging

        /// <summary>
        /// Initializes trace logging by redirecting Trace output to a log file.
        /// This method ensures the log file is opened in append mode with proper flushing
        /// and listener setup for reliable logging output.
        /// </summary>
        /// <param name="logFilePath">The full path to the log file.</param>
        private void InitializeTraceLogging(string logFilePath)
        {
            // Close any existing listeners before setting up a new one.
            CloseTraceLogging();

            // Open the log file stream in append mode, allowing concurrent read/write access.
            var logFileStream = new FileStream(
                logFilePath,
                FileMode.Append,      // Ensure new logs are added to the end of the file.
                FileAccess.Write,     // Only write access is needed.
                FileShare.ReadWrite); // Allow other processes to read or write concurrently.

            // Wrap the file stream with a StreamWriter for better buffering and encoding control.
            var writer = new StreamWriter(logFileStream)
            {
                AutoFlush = true // Ensure log messages are written immediately without buffering delay.
            };

            // Create a trace listener that writes to the StreamWriter.
            var listener = new FilteredTraceListener(writer)
            {
                // Include timestamp in trace output.
                TraceOutputOptions = TraceOptions.DateTime
            };

            // Enable automatic flushing of the Trace buffer after each write.
            Trace.AutoFlush = true;

            // Add our configured listener to the Trace system.
            Trace.Listeners.Add(listener);
        }

        /// <summary>
        /// Flushes and closes all trace listeners, then clears them from the Trace system.
        /// Call this during application shutdown or before reinitializing logging.
        /// </summary>
        private static void CloseTraceLogging()
        {
            foreach (TraceListener listener in Trace.Listeners)
            {
                listener.Flush(); // Ensure all buffered output is written to the underlying stream.
                listener.Close(); // Release the file handle and resources used by the listener.
            }

            Trace.Listeners.Clear(); // Remove all listeners from the Trace system.
        }

        #endregion Debug Logging

        #region Methods

        /// <summary>
        /// Update the fields and buttons in the form.
        /// </summary>
        private void UpdateFormControls()
        {
            UpdateFormFields();

            // Check if the run button is now enabled/disabled.
            _dockPane.CheckRunButton();
        }

        /// <summary>
        /// Update the fields in the form.
        /// </summary>
        private void UpdateFormFields()
        {
            OnPropertyChanged(nameof(SearchRefText));
            OnPropertyChanged(nameof(SiteNameText));
            OnPropertyChanged(nameof(SiteNameTextVisibility));
            OnPropertyChanged(nameof(OrganisationText));
            OnPropertyChanged(nameof(OrganisationTextVisibility));
            OnPropertyChanged(nameof(MapLayersList));
            OnPropertyChanged(nameof(LayersListEnabled));
            OnPropertyChanged(nameof(BufferSizeText));
            OnPropertyChanged(nameof(BufferUnitsList));
            OnPropertyChanged(nameof(BufferUnitsListEnabled));
            OnPropertyChanged(nameof(SelectedBufferUnitsIndex));
            OnPropertyChanged(nameof(AddToMapList));
            OnPropertyChanged(nameof(AddToMapListEnabled));
            OnPropertyChanged(nameof(SelectedAddToMap));
            OnPropertyChanged(nameof(AddToMapListVisibility));
            OnPropertyChanged(nameof(OverwriteLabelsList));
            OnPropertyChanged(nameof(OverwriteLabelsListEnabled));
            OnPropertyChanged(nameof(SelectedOverwriteLabels));
            OnPropertyChanged(nameof(OverwriteLabelsListVisibility));
            OnPropertyChanged(nameof(CombinedSitesList));
            OnPropertyChanged(nameof(CombinedSitesListEnabled));
            OnPropertyChanged(nameof(SelectedCombinedSites));
            OnPropertyChanged(nameof(CombinedSitesListVisibility));
            OnPropertyChanged(nameof(Message));
        }

        /// <summary>
        /// Lookup the search reference in the database.
        /// </summary>
        /// <param name="searchRefText"></param>
        /// <param name="siteName"></param>
        /// <param name="organisation"></param>
        /// <returns>bool</returns>
        public bool LookupSearchRef(string searchRefText, ref string siteName, ref string organisation)
        {
            // Use connection string for .accdb or .mdb as appropriate.
            string connectionString;
            if (FileFunctions.GetExtension(_databasePath).Equals(".accdb", StringComparison.OrdinalIgnoreCase))
                connectionString = "Provider='Microsoft.ACE.OLEDB.12.0';data source='" + _databasePath + "'";
            else
                connectionString = "Provider='Microsoft.Jet.OLEDB.4.0';data source='" + _databasePath + "'";

            // Build the list of columns to retrieve.
            string mapColumns = _databaseSiteColumn;
            if (!string.IsNullOrEmpty(_databaseOrgColumn))
                mapColumns += "," + _databaseOrgColumn;

            // Build the search query.
            string searchQuery = "SELECT " + mapColumns + " FROM " + _databaseTable + " WHERE LCASE(" + _databaseRefColumn + ") = " + '"' + searchRefText + '"';

            siteName = null;
            organisation = null;

            // Connect to the database.
            OleDbConnection oleDbConnection;
            try
            {
                oleDbConnection = new(connectionString);
            }
            catch (Exception)
            {
                FileFunctions.WriteLine(_logFile, "Error creating a database connection");
                return false;
            }

            DataSet dataSet = new();
            try
            {
                // Create an adapter.
                OleDbCommand oleDbCommand = new(searchQuery, oleDbConnection);
                OleDbDataAdapter dataAdapter = new(oleDbCommand);

                // Open the database.
                oleDbConnection.Open();

                // File the adapter for the table.
                dataAdapter.Fill(dataSet, _databaseTable);

                dataAdapter.Dispose();
                oleDbCommand.Dispose();
            }
            catch (Exception)
            {
                FileFunctions.WriteLine(_logFile, "Error Retrieving the required data from the database");
                return false;
            }
            finally
            {
                oleDbConnection.Dispose();
                oleDbConnection.Close();
            }

            try
            {
                DataRowCollection dataRowCollection = dataSet.Tables[_databaseTable].Rows;
                foreach (DataRow aRow in dataRowCollection) // Really there should only be one. We can check for this.
                {
                    // Get the site name and organisation names.
                    siteName = aRow[0].ToString();
                    if (!string.IsNullOrEmpty(_databaseOrgColumn))
                        organisation = aRow[1].ToString();
                }
            }
            catch
            {
                FileFunctions.WriteLine(_logFile, "Error Finding the required table in the table");
                return false;
            }
            finally
            {
                dataSet.Dispose();
            }

            return true;
        }

        /// <summary>
        /// Set all of the form fields to their default values.
        /// </summary>
        /// <param name="reset"></param>
        /// <returns></returns>
        public async Task ResetFormAsync()
        {
            // Clear the selections first (to avoid selections being retained).
            if (_mapLayersList != null)
            {
                foreach (MapLayer layer in _mapLayersList)
                {
                    layer.IsSelected = false;
                }
            }

            // Search ref and site name.
            SearchRefText = null;
            SiteNameText = null;
            OrganisationText = null;

            // Buffer size and units.
            BufferSizeText = _toolConfig.DefaultBufferSize.ToString();
            BufferUnitsList = _bufferUnitOptionsDisplay;
            if (_toolConfig.DefaultBufferUnit > 0)
                SelectedBufferUnitsIndex = _toolConfig.DefaultBufferUnit - 1;

            // Keep selected map layers.
            KeepSelectedLayers = _toolConfig.DefaultKeepSelectedLayers != null && (bool)_toolConfig.DefaultKeepSelectedLayers;

            // Add layers to map.
            SelectedAddToMap = _toolConfig.DefaultAddSelectedLayers < 1 ? AddToMapList[0] : AddToMapList[_toolConfig.DefaultAddSelectedLayers - 1];

            // Overwrite map layers.
            SelectedOverwriteLabels = _toolConfig.DefaultOverwriteLabels < 1 ? OverwriteLabelsList[0] : OverwriteLabelsList[_toolConfig.DefaultOverwriteLabels - 1];

            // Combined sites table.
            SelectedCombinedSites = _toolConfig.DefaultCombinedSitesTable < 1 ? CombinedSitesList[0] : CombinedSitesList[_toolConfig.DefaultCombinedSitesTable - 1];

            // Log file.
            ClearLogFile = _toolConfig.DefaultClearLogFile;
            OpenLogFile = _toolConfig.DefaultOpenLogFile;

            // Pause map.
            PauseMap = _toolConfig.PauseMap;

            // Reload the form layers.
            await LoadLayersAsync(false);
        }

        /// <summary>
        /// Load the list of open GIS layers.
        /// </summary>
        /// <param name="reset"></param>
        /// <param name="message"></param>
        /// <returns></returns>
        public async Task LoadLayersAsync(bool message)
        {
            // If already processing then exit.
            if (_dockPane.ProcessStatus != null)
                return;

            // Expand the lists (ready to be resized later).
            _mapLayersListHeight = null;

            _dockPane.FormLoading = true;
            _dockPane.ProgressUpdate("Loading form...");

            // Force UI update
            //await Task.Delay(1);
            await Application.Current.Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Background);

            // Clear any messages.
            ClearMessage();

            // Update the fields and buttons in the form.
            UpdateFormControls();

            // Reload the list of GIS map layers (don't wait).
            string loadMapResult = await LoadMapLayersAsync();

            // Set the list of open layers.
            MapLayersList = new ObservableCollection<MapLayer>(_openMapLayersList);

            // Hide progress update.
            _dockPane.ProgressUpdate(null, -1, -1);

            // Show a message if there are no open layers.
            if (!_mapLayersList.Any())
                ShowMessage("No search layers in active map.", MessageType.Warning);

            // Indicate the form has loaded.
            _dockPane.FormLoading = false;

            // Update the fields and buttons in the form.
            UpdateFormControls();

            // Force list column widths to reset.
            MapLayersListExpandCommandClick(null);

            // Show any message from loading the map layers list.
            if (loadMapResult != null!)
            {
                ShowMessage(loadMapResult, MessageType.Warning);
                if (message)
                    MessageBox.Show(loadMapResult, _displayName, MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
        }

        /// <summary>
        /// Load the list of open GIS layers.
        /// </summary>
        /// <returns></returns>
        public async Task<string> LoadMapLayersAsync()
        {
            // Reset the list of open layers.
            _openMapLayersList = [];

            // Reset the list of closed layers.
            _closedMapLayersList = [];

            // Load the map layer variables from the XML profile.
            bool mapLayersloaded = false;
            try
            {
                await Task.Run(() =>
                {
                    if (_toolConfig.GetMapVariables())
                        mapLayersloaded = true;
                });
            }
            catch (Exception ex)
            {
                // Only report message if user was prompted for the XML
                // file (i.e. the user interface has already loaded).
                return "Error loading XML file. " + ex.Message;
            }

            if (!mapLayersloaded)
                return "Error loading Map variables from XML file.";

            // Get all of the map layer details.
            _mapLayers = _toolConfig.MapLayers;

            await Task.Run(async () =>
            {
                if (_mapFunctions == null || _mapFunctions.MapName == null || MapView.Active is null || MapView.Active.Map.Name != _mapFunctions.MapName)
                {
                    // Create a new map functions object.
                    _mapFunctions = new();
                }

                // Check if there is an active map.
                bool mapOpen = _mapFunctions.MapName != null;

                if (mapOpen)
                {
                    List<MapLayer> allLayers = _mapLayers;

                    // Loop through all of the layers to check if they are open
                    // in the active map.
                    foreach (MapLayer layer in allLayers)
                    {
                        if (await _mapFunctions.FindLayerAsync(layer.LayerName, null) != null)
                        {
                            // Preselect layer if required.
                            layer.IsSelected = layer.PreselectLayer;

                            // Add the open layers to the list.
                            _openMapLayersList.Add(layer);
                        }
                        else
                        {
                            // Only add if the user wants to be warned of this one.
                            if (layer.LoadWarning)
                                _closedMapLayersList.Add(layer.LayerName);
                        }
                    }
                }
            });

            // Show a message if there are no open map layers.
            if (_openMapLayersList.Count == 0)
                return "No map layers in active map.";

            // Warn the user of any closed map layers.
            int closedLayerCount = _closedMapLayersList.Count;
            if (closedLayerCount > 0)
            {
                string closedLayerWarning = "";
                if (closedLayerCount == 1)
                {
                    closedLayerWarning = "Layer '" + _closedMapLayersList[0] + "' is not loaded.";
                }
                else
                {
                    closedLayerWarning = string.Format("{0} map layers are not loaded.", closedLayerCount.ToString());
                }

                return closedLayerWarning;
            }

            return null;
        }

        /// <summary>
        /// Clear the list of open GIS layers.
        /// </summary>
        /// <param name="selectedTable"></param>
        /// <returns></returns>
        public void ClearLayers()
        {
            // Clear the list of open layers.
            _mapLayersList = [];

            // Update the fields and buttons in the form.
            UpdateFormControls();
        }

        /// <summary>
        /// Validate and run the search.
        /// </summary>
        private async Task<bool> RunSearchAsync()
        {
            if (_mapFunctions == null || _mapFunctions.MapName == null || MapView.Active.Map.Name != _mapFunctions.MapName)
            {
                // Create a new map functions object.
                _mapFunctions = new();
            }

            // Reset search errors flag.
            _searchErrors = false;

            // Save the parameters.
            string searchRef = SearchRefText;
            string siteName = SiteNameText;
            string organisation = OrganisationText;
            string bufferSize = BufferSizeText;
            int bufferUnitIndex = SelectedBufferUnitsIndex;

            // Selected layers.
            _selectedLayers = MapLayersList.Where(p => p.IsSelected).ToList();

            // List of output map and window names.
            List<string> mapWindowNames = _toolConfig.MapNames;
            List<string> layoutWindowNames = _toolConfig.LayoutNames;

            string searchRefElement = _toolConfig.SearchRefElement;
            string siteNameElement = _toolConfig.SiteNameElement;
            string organisationElement = _toolConfig.OrganisationElement;
            string radiusElement = _toolConfig.RadiusElement;
            string bespokeElements = _toolConfig.BespokeElements;

            // Zoom ratio and list of zoom scales for the layout windows.
            double zoomRatio = _toolConfig.ZoomRatio;
            List<int> zoomScales = _toolConfig.ZoomScales;

            // What is the selected buffer unit?
            string bufferUnitText = _bufferUnitOptionsDisplay[bufferUnitIndex]; // Unit to be used in reporting.
            string bufferUnitProcess = _bufferUnitOptionsProcess[bufferUnitIndex]; // Unit to be used in process (because of American spelling).
            string bufferUnitShort = _bufferUnitOptionsShort[bufferUnitIndex]; // Unit to be used in file naming (abbreviation).

            // Should labels be converted to annocation?
            bool convertLabelsToAnnotation = _toolConfig.ConvertLabelsToAnnotation;

            // What is the area measurement unit?
            string areaMeasureUnit = _toolConfig.AreaMeasurementUnit;

            // Will the selected layers be kept?
            bool keepSelectedLayers = KeepSelectedLayers;

            // Will the selected layers be added to the map with labels?
            AddSelectedLayersOptions addSelectedLayersOption = AddSelectedLayersOptions.No;
            if (_defaultAddSelectedLayers > 0)
            {
                if (SelectedAddToMap.Equals("no", StringComparison.OrdinalIgnoreCase))
                    addSelectedLayersOption = AddSelectedLayersOptions.No;
                else if (SelectedAddToMap.Contains("with labels", StringComparison.OrdinalIgnoreCase))
                    addSelectedLayersOption = AddSelectedLayersOptions.WithLabels;
                else if (SelectedAddToMap.Contains("without labels", StringComparison.OrdinalIgnoreCase))
                    addSelectedLayersOption = AddSelectedLayersOptions.WithoutLabels;
            }

            // Will the labels on map layers be overwritten?
            OverwriteLabelOptions overwriteLabelOption = OverwriteLabelOptions.No;
            if (_defaultOverwriteLabels > 0)
            {
                if (SelectedOverwriteLabels.Equals("no", StringComparison.OrdinalIgnoreCase))
                    overwriteLabelOption = OverwriteLabelOptions.No;
                else if (SelectedOverwriteLabels.Contains("reset each layer", StringComparison.OrdinalIgnoreCase))
                    overwriteLabelOption = OverwriteLabelOptions.ResetByLayer;
                else if (SelectedOverwriteLabels.Contains("reset each group", StringComparison.OrdinalIgnoreCase))
                    overwriteLabelOption = OverwriteLabelOptions.ResetByGroup;
                else if (SelectedOverwriteLabels.Contains("do not reset", StringComparison.OrdinalIgnoreCase))
                    overwriteLabelOption = OverwriteLabelOptions.DoNotReset;
            }

            // Will the combined sites table be created and overwritten?
            CombinedSitesTableOptions combinedSitesTableOption = CombinedSitesTableOptions.None;
            if (_defaultCombinedSitesTable > 0)
            {
                if (SelectedCombinedSites.Equals("none", StringComparison.OrdinalIgnoreCase))
                    combinedSitesTableOption = CombinedSitesTableOptions.None;
                else if (SelectedCombinedSites.Contains("append", StringComparison.OrdinalIgnoreCase))
                    combinedSitesTableOption = CombinedSitesTableOptions.Append;
                else if (SelectedCombinedSites.Contains("overwrite", StringComparison.OrdinalIgnoreCase))
                    combinedSitesTableOption = CombinedSitesTableOptions.Overwrite;
            }

            // Fix any illegal characters in the site name string.
            siteName = StringFunctions.StripIllegals(siteName, _repChar);

            // Create the ref string from the search reference.
            string reference = searchRef.Replace("/", _repChar);

            // Create the shortref from the search reference by
            // getting rid of any characters.
            string shortRef = StringFunctions.KeepNumbersAndSpaces(reference, _repChar);

            // Find the subref part of this reference.
            string subref = StringFunctions.GetSubref(shortRef, _repChar);

            // Create the radius from the buffer size and units
            string radius = bufferSize + bufferUnitShort;

            // Replace any standard strings in the variables.
            _saveFolder = StringFunctions.ReplaceSearchStrings(_toolConfig.SaveFolder, reference, siteName, shortRef, subref, radius, organisation);
            _extractFolder = StringFunctions.ReplaceSearchStrings(_toolConfig.ExtractFolder, reference, siteName, shortRef, subref, radius, organisation);
            _gisFolder = StringFunctions.ReplaceSearchStrings(_toolConfig.GISFolder, reference, siteName, shortRef, subref, radius, organisation);
            _logFileName = StringFunctions.ReplaceSearchStrings(_toolConfig.LogFileName, reference, siteName, shortRef, subref, radius, organisation);
            _combinedSitesTableName = StringFunctions.ReplaceSearchStrings(_toolConfig.CombinedSitesTableName, reference, siteName, shortRef, subref, radius, organisation);
            _bufferPrefix = StringFunctions.ReplaceSearchStrings(_toolConfig.BufferPrefix, reference, siteName, shortRef, subref, radius, organisation);
            _searchLayerName = StringFunctions.ReplaceSearchStrings(_toolConfig.SearchOutputName, reference, siteName, shortRef, subref, radius, organisation);
            _groupLayerName = StringFunctions.ReplaceSearchStrings(_toolConfig.GroupLayerName, reference, siteName, shortRef, subref, radius, organisation);

            // Replace any standard strings in the bespoke elements.
            bespokeElements = StringFunctions.ReplaceSearchStrings(bespokeElements, searchRef, siteName, shortRef, subref, radius, organisation);

            // Set the date variables.
            DateTime dateNow = DateTime.Now;
            string dateDD = dateNow.ToString("dd");
            string dateMM = dateNow.ToString("MM");
            string dateMMM = dateNow.ToString("MMM");
            string dateMMMM = dateNow.ToString("MMMM");
            string dateYY = dateNow.ToString("yy");
            double dateQtr = (Math.Ceiling(dateNow.Month / 3.0 + 2) % 4) + 1;
            string dateQQ = dateQtr.ToString("00");
            string dateYYYY = dateNow.ToString("yyyy");
            string dateFFFF = StringFunctions.FinancialYear(dateNow);

            // Replace any date variables in the bespoke elements.
            bespokeElements = bespokeElements.Replace("%dd%", dateDD).Replace("%mm%", dateMM).Replace("%mmm%", dateMMM).Replace("%mmmm%", dateMMMM);
            bespokeElements = bespokeElements.Replace("%yy%", dateYY).Replace("%qq%", dateQQ).Replace("%yyyy%", dateYYYY).Replace("%ffff%", dateFFFF);

            // Replace any date variables in the file and folder variables.
            _saveFolder = _saveFolder.Replace("%dd%", dateDD).Replace("%mm%", dateMM).Replace("%mmm%", dateMMM).Replace("%mmmm%", dateMMMM);
            _saveFolder = _saveFolder.Replace("%yy%", dateYY).Replace("%qq%", dateQQ).Replace("%yyyy%", dateYYYY).Replace("%ffff%", dateFFFF);
            _extractFolder = _extractFolder.Replace("%dd%", dateDD).Replace("%mm%", dateMM).Replace("%mmm%", dateMMM).Replace("%mmmm%", dateMMMM);
            _extractFolder = _extractFolder.Replace("%yy%", dateYY).Replace("%qq%", dateQQ).Replace("%yyyy%", dateYYYY).Replace("%ffff%", dateFFFF);
            _gisFolder = _gisFolder.Replace("%dd%", dateDD).Replace("%mm%", dateMM).Replace("%mmm%", dateMMM).Replace("%mmmm%", dateMMMM);
            _gisFolder = _gisFolder.Replace("%yy%", dateYY).Replace("%qq%", dateQQ).Replace("%yyyy%", dateYYYY).Replace("%ffff%", dateFFFF);
            _logFileName = _logFileName.Replace("%dd%", dateDD).Replace("%mm%", dateMM).Replace("%mmm%", dateMMM).Replace("%mmmm%", dateMMMM);
            _logFileName = _logFileName.Replace("%yy%", dateYY).Replace("%qq%", dateQQ).Replace("%yyyy%", dateYYYY).Replace("%ffff%", dateFFFF);

            // Split the bespoke elements into pairs of element names and contents.
            List<string> bespokeElementNames = [];
            List<string> bespokeContents = [];
            string[] pairs = bespokeElements.Split('$');

            foreach (string pair in pairs)
            {
                // Split each pair into element and content
                string[] parts = pair.Split(';');
                if (parts.Length == 2)
                {
                    bespokeElementNames.Add(parts[0]);
                    bespokeContents.Add(parts[1]);
                }
            }

            // Remove any illegal characters from the names.
            _saveFolder = StringFunctions.StripIllegals(_saveFolder, _repChar);
            _extractFolder = StringFunctions.StripIllegals(_extractFolder, _repChar);
            _gisFolder = StringFunctions.StripIllegals(_gisFolder, _repChar);
            _logFileName = StringFunctions.StripIllegals(_logFileName, _repChar, true);
            _combinedSitesTableName = StringFunctions.StripIllegals(_combinedSitesTableName, _repChar);
            _bufferPrefix = StringFunctions.StripIllegals(_bufferPrefix, _repChar);
            _searchLayerName = StringFunctions.StripIllegals(_searchLayerName, _repChar);
            _groupLayerName = StringFunctions.StripIllegals(_groupLayerName, _repChar);

            // Trim any trailing spaces (directory functions don't deal with them well).
            _saveFolder = _saveFolder.Trim();

            // Clear the GIS folder if the layers are not to be kept so an
            // empty output folder isn't created.
            if (!keepSelectedLayers)
                _gisFolder = null;

            // Create output folders if required.
            if (!CreateOutputFolders(_saveRootDir, _saveFolder, _extractFolder, _gisFolder, ref _outputPath, ref _gisPath))
            {
                MessageBox.Show("Cannot create output folders");
                return false;
            }

            // Create log file (if necessary).
            _logFile = _outputPath + @"\" + _logFileName;
            if (FileFunctions.FileExists(_logFile) && ClearLogFile)
            {
                try
                {
                    File.Delete(_logFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot clear log file " + _logFile + ". Please make sure this file is not open in another window. " +
                        "System error: " + ex.Message);
                    return false;
                }
            }

            // Initialize the application-wide trace logging to redirect the output to the log file.
            InitializeTraceLogging(_logFile);

            // Replace any illegal characters in the user name string.
            _userID = StringFunctions.StripIllegals(Environment.UserName, "_", false);

            // Replace hyphen with underscore.
            _userID = _userID.Replace('-', '_');

            // User ID should be something at least.
            if (string.IsNullOrEmpty(_userID))
            {
                _userID = "Temp";
                FileFunctions.WriteLine(_logFile, "User ID not found. User ID used will be 'Temp'");
            }

            // Count the number of layers to process.
            int stepsMax = SelectedLayers.Count;
            int stepNum = 0;

            // Build a list of open map windows.
            List<Map> mapWindows = [];
            foreach (string mapName in mapWindowNames)
            {
                // Get the map pane with the window (caption) name.
                Map map = await _mapFunctions.GetMapFromCaptionAsync(mapName);

                if (map != null)
                    mapWindows.Add(map);
            }

            // Build a list of open layout windows.
            List<Layout> layoutWindows = [];
            foreach (string layoutName in layoutWindowNames)
            {
                // Get the map pane with the window (caption) name.
                Layout layout = await _mapFunctions.GetLayoutFromCaptionAsync(layoutName);

                if (layout != null)
                    layoutWindows.Add(layout);
            }

            // Re-activate the original map window.
            if (await _mapFunctions.ActivateMapAsync(null) == null)
            {
                FileFunctions.WriteLine(_logFile, "Error re-activating active map");
                _searchErrors = true;

                return false;
            }

            // Indicate the search is running.
            _dockPane.SearchStatus = DockpaneMainViewModel.SearchStatuses.Running;

            // Write the first line to the log file.
            FileFunctions.WriteLine(_logFile, "-----------------------------------------------------------------------");
            FileFunctions.WriteLine(_logFile, "Processing search '" + searchRef + "'");
            FileFunctions.WriteLine(_logFile, "-----------------------------------------------------------------------");

            FileFunctions.WriteLine(_logFile, "Parameters are as follows:");
            FileFunctions.WriteLine(_logFile, "Buffer distance: " + radius);
            FileFunctions.WriteLine(_logFile, "Output location: " + _saveRootDir + @"\" + _saveFolder);
            FileFunctions.WriteLine(_logFile, "Layers to process: " + SelectedLayers.Count.ToString());
            FileFunctions.WriteLine(_logFile, "Area measurement unit: " + areaMeasureUnit);

            // Create the search query.
            string searchClause = _searchColumn + " = '" + searchRef + "'";

            _dockPane.ProgressUpdate("Selecting feature(s)...", 0);
            //stepNum += 1;

            // Count the features matching the search reference.
            if (await CountSearchFeaturesAsync(searchClause) == 0)
            {
                _searchErrors = true;
                return false;
            }

            // Prepare the temporary geodatabase
            if (!await PrepareTemporaryGDBAsync())
            {
                _searchErrors = true;
                return false;
            }

            // Stop if the user cancelled the process.
            if (_dockPane.SearchStatus == DockpaneMainViewModel.SearchStatuses.Cancelled)
                return false;

            // Pause the active map redrawing.
            if (PauseMap)
                _mapFunctions.PauseDrawing(true, null);

            // Select the feature matching the search reference in the map.
            if (!await _mapFunctions.SelectLayerByAttributesAsync(_inputLayerName, searchClause, SelectionCombinationMethod.New))
            {
                _searchErrors = true;
                return false;
            }

            // Update the table in the active map if required.
            if (_updateTable && (!string.IsNullOrEmpty(_siteColumn) || !string.IsNullOrEmpty(_orgColumn) || !string.IsNullOrEmpty(_radiusColumn)))
            {
                FileFunctions.WriteLine(_logFile, "Updating attributes in search layer ...");

                if (!await _mapFunctions.UpdateFeaturesAsync(_inputLayerName, _siteColumn, siteName, _orgColumn, organisation, _radiusColumn, radius, null))
                {
                    FileFunctions.WriteLine(_logFile, "Error updating attributes in search layer");
                    _searchErrors = true;
                    return false;
                }
            }

            // The output file for the search features is a shapefile in the root save directory.
            _searchOutputFile = _outputPath + "\\" + _searchLayerName + ".shp";

            // Remove the search feature layer from the search map
            // in case there is one already there from a different folder.
            await _mapFunctions.RemoveLayerAsync(_searchLayerName, null);

            // Remove the search feature layer from all other map windows
            // in case there is one already there from a different folder.
            foreach (Map map in mapWindows)
                await _mapFunctions.RemoveLayerAsync(_searchLayerName, map);

            // Save the selected feature(s).
            if (!await SaveSearchFeaturesAsync())
            {
                _searchErrors = true;
                return false;
            }

            // Stop if the user cancelled the process.
            if (_dockPane.SearchStatus == DockpaneMainViewModel.SearchStatuses.Cancelled)
                return false;

            _dockPane.ProgressUpdate("Buffering feature(s)...", 0);
            //stepNum += 1;

            // Set the buffer layer name by appending the radius.
            _bufferLayerName = _bufferPrefix + "_" + radius;
            if (_bufferLayerName.Contains('.'))
                _bufferLayerName = _bufferLayerName.Replace('.', '_');

            // The output file for the buffer is a shapefile in the root save directory.
            _bufferOutputFile = _outputPath + "\\" + _bufferLayerName + ".shp";

            // Remove the buffer layer from the search map
            // in case there is one already there from a different folder.
            await _mapFunctions.RemoveLayerAsync(_bufferLayerName, null);

            // Remove the buffer layer from all other map windows
            // in case there is one already there from a different folder.
            foreach (Map map in mapWindows)
                await _mapFunctions.RemoveLayerAsync(_bufferLayerName, map);

            // Buffer search feature(s).
            if (!await BufferSearchFeaturesAsync(bufferSize, bufferUnitProcess, bufferUnitShort))
            {
                _searchErrors = true;
                return false;
            }

            // Find the buffer layer by name in the active map.
            _bufferLayer = await _mapFunctions.FindLayerAsync(_bufferLayerName, null);
            if (_bufferLayer == null)
            {
                FileFunctions.WriteLine(_logFile, "Error finding buffer layer '" + _bufferLayerName + "' in active map");
                _searchErrors = true;

                return false;
            }

            // Get the full layer path (in case it's nested in one or more groups).
            _bufferLayerPath = await _mapFunctions.GetLayerPathAsync(_bufferLayerName);

            // Change the combined sites table option if the name is empty.
            if (string.IsNullOrWhiteSpace(_combinedSitesTableName))
                combinedSitesTableOption = CombinedSitesTableOptions.None;

            // Start the combined sites table before we do any analysis.
            _combinedSitesOutputFile = _outputPath + @"\" + _combinedSitesTableName + "." + _combinedSitesTableFormat;
            if (!CreateCombinedSitesTable(_combinedSitesOutputFile, combinedSitesTableOption))
            {
                _searchErrors = true;
                return false;
            }

            // Get any groups and initialise required layers.
            if (overwriteLabelOption == OverwriteLabelOptions.ResetByLayer ||
                overwriteLabelOption == OverwriteLabelOptions.ResetByGroup)
            {
                _mapGroupNames = _selectedLayers.Select(l => l.NodeGroup).Distinct().ToList();
                _mapGroupLabels = [];
                foreach (string groupName in _mapGroupNames)
                {
                    // Each group has its own label counter.
                    _mapGroupLabels.Add(1);
                }
            }

            // Keep track of the label numbers.
            _maxLabel = 1;

            bool success;

            int layerNum = 0;
            int layerCount = SelectedLayers.Count;
            foreach (MapLayer selectedLayer in SelectedLayers)
            {
                // Stop if the user cancelled the process.
                if (_dockPane.SearchStatus == DockpaneMainViewModel.SearchStatuses.Cancelled)
                    break;

                // Get the layer name.
                string mapNodeGroup = selectedLayer.NodeGroup;
                string mapNodeLayer = selectedLayer.NodeLayer;

                _dockPane.ProgressUpdate("Processing '" + mapNodeGroup + " - " + mapNodeLayer + "'...", stepNum, stepsMax);
                stepNum += 1;

                layerNum += 1;
                FileFunctions.WriteLine(_logFile, "");
                FileFunctions.WriteLine(_logFile, "Starting analysis for '" + selectedLayer.NodeName + "' (" + layerNum + " of " + layerCount + ")");

                // Loop through the map layers, processing each one.
                success = await ProcessMapLayerAsync(selectedLayer, reference, siteName, shortRef, subref, radius, organisation, areaMeasureUnit, keepSelectedLayers, addSelectedLayersOption, overwriteLabelOption, combinedSitesTableOption);

                // Keep track of any errors.
                if (!success)
                    _searchErrors = true;
            }

            _dockPane.ProgressUpdate("Finishing up...", stepNum, 0);

            // Clean up after the search.
            await CleanUpSearchAsync();

            // If there were errors at this stage then exit.
            if (_searchErrors)
                return false;

            // If the process was cancelled when exit.
            if (_dockPane.SearchStatus == DockpaneMainViewModel.SearchStatuses.Cancelled)
                return false;

            // Set the search reference, site name, organisation and radius in all of the layouts.
            if (!await _mapFunctions.UpdateLayoutsTextAsync(layoutWindowNames, searchRefElement, searchRef, siteNameElement, siteName,
                organisationElement, organisation, radiusElement, radius, bespokeElementNames, bespokeContents))
            {
                _searchErrors = true;

                return false;
            }

            string bufferLayerName = _bufferLayerName;

            // If the buffer is being kept.
            if (_keepBuffer)
            {
                // Add the buffer layer to all of the map windows.
                bufferLayerName = await AddBufferLayerToMapsAsync(mapWindows);
                if (bufferLayerName == null)
                {
                    _searchErrors = true;
                    return false;
                }
            }
            else
            {
                // Remove the buffer layer from all maps.
                await RemoveLayerFromAllMapsAsync(bufferLayerName, _bufferOutputFile, mapWindows);
            }

            string searchLayerName = _searchLayerName;

            // If the search area is being kept.
            if (_keepSearchFeatureExtensions.Contains(_searchLayerExtension))
            {
                // Add the search area layer to all of the map windows.
                searchLayerName = await AddSearchLayerToMapsAsync(mapWindows);
                if (searchLayerName == null)
                {
                    _searchErrors = true;
                    return false;
                }
            }
            else
            {
                string searchOutputFile = _outputPath + "\\" + _searchLayerName + ".shp";

                // Remove the search feature layer from all maps.
                await RemoveLayerFromAllMapsAsync(searchLayerName, searchOutputFile, mapWindows);
            }

            // Indicate the search is completing.
            _dockPane.SearchStatus = DockpaneMainViewModel.SearchStatuses.Completing;

            // Zoom to the buffer layer extent (or the search layer extent if there is no buffer).
            string targetLayer = bufferSize == "0" ? searchLayerName : bufferLayerName;

            // Zoom in to the search layer map.
            if (!await _mapFunctions.ZoomToLayerInMapAsync(targetLayer, false, zoomRatio, null, null))
            {
                FileFunctions.WriteLine(_logFile, "Error zooming in active map");
                _searchErrors = true;

                return false;
            }

            // Zoom in to all other map windows.
            foreach (Map map in mapWindows)
            {
                // Activate the map window.
                if (await _mapFunctions.ActivateMapAsync(map) == null)
                {
                    FileFunctions.WriteLine(_logFile, $"Error activating map: {map.Name}");
                    _searchErrors = true;

                    return false;
                }

                // Zoom to the layer in the map window.
                if (!await _mapFunctions.ZoomToLayerInMapAsync(targetLayer, false, zoomRatio, null, map))
                {
                    FileFunctions.WriteLine(_logFile, $"Error zooming in map: {map.Name}");
                    _searchErrors = true;

                    return false;
                }

                // Convert labels to annotation if required.
                if (convertLabelsToAnnotation)
                {
                    // Create the annotation file geodatabase if it doesn't exist.
                    string annoGDBName = _gisPath + @"\Anno.gdb";
                    Geodatabase annoGDB = null;
                    if (!FileFunctions.DirExists(annoGDBName))
                    {
                        annoGDB = ArcGISFunctions.CreateFileGeodatabase(annoGDBName);
                        if (annoGDB == null)
                        {
                            FileFunctions.WriteLine(_logFile, "Error creating annotation geodatabase " + annoGDBName);
                            _searchErrors = true;

                            return false;
                        }

                        FileFunctions.WriteLine(_logFile, "Annotation geodatabase created");
                    }

                    // Create the annotation group layer name.
                    string baseAnnoGroupLayerName = $"{_groupLayerName}_Anno";
                    string annoGroupLayerName = await GetNextAnnoGroupLayerNameAsync(baseAnnoGroupLayerName, map);

                    if (!await _mapFunctions.ConvertLabelsToAnnotationAsync(map.Name, "", annoGDBName, "_Anno", "DISPLAY", null, true, "", false, annoGroupLayerName))
                    {
                        FileFunctions.WriteLine(_logFile, $"Error converting all labels to annotation in map: {map.Name}");
                        _searchErrors = true;

                        return false;
                    }

                    // Remove the annotation group layer from the map if it is empty.
                    await _mapFunctions.RemoveGroupLayerAsync(annoGroupLayerName, map);
                }
            }

            // Zoom in to all layout windows.
            foreach (Layout layout in layoutWindows)
            {
                // Activate the layout window.
                if (await _mapFunctions.ActivateLayoutAsync(layout) == null)
                {
                    FileFunctions.WriteLine(_logFile, $"Error activating layout: {layout.Name}");
                    _searchErrors = true;

                    return false;
                }

                if (!await _mapFunctions.ZoomToLayerInLayoutAsync(layout, targetLayer, selectedOnly: false,
                    ratio: zoomRatio, scale: null, validScales: zoomScales, mapFrameName: "Map Frame"))
                {
                    FileFunctions.WriteLine(_logFile, $"Error zooming in layout: {layout}");
                    _searchErrors = true;

                    return false;
                }
            }

            // Re-activate the original map window.
            if (await _mapFunctions.ActivateMapAsync(null) == null)
            {
                FileFunctions.WriteLine(_logFile, "Error re-activating active map");
                _searchErrors = true;

                return false;
            }

            return true;
        }

        /// <summary>
        /// Gets the next available annotation group layer name by appending a numeric suffix (without underscore),
        /// starting with the base name (e.g., "MyLayer_Anno", "MyLayer_Anno2", "MyLayer_Anno3", ...).
        /// </summary>
        /// <param name="baseGroupLayerName">The base name to use (e.g. "MyLayer_Anno").</param>
        /// <param name="map">The map in which to check for name conflicts.</param>
        /// <returns>A unique group layer name that does not already exist in the map.</returns>
        internal async Task<string> GetNextAnnoGroupLayerNameAsync(string baseGroupLayerName, Map map)
        {
            if (string.IsNullOrEmpty(baseGroupLayerName) || map == null)
                return null;

            string nextGroupLayerName = baseGroupLayerName;
            int suffix = 2;

            // First try the base name as-is.
            GroupLayer existingLayer = await _mapFunctions.FindGroupLayerAsync(nextGroupLayerName, map);
            if (existingLayer == null)
                return nextGroupLayerName;

            // Try increasing numeric suffix.
            while (true)
            {
                nextGroupLayerName = $"{baseGroupLayerName}{suffix}";
                if (await _mapFunctions.FindGroupLayerAsync(nextGroupLayerName, map) == null)
                    return nextGroupLayerName;

                suffix++;
            }
        }

        /// <summary>
        /// Indicate that the search process has stopped (either
        /// successfully or otherwise).
        /// </summary>
        /// <param name="searchRef"></param>
        /// <param name="success"></param>
        private void StopSearch(string message, string image)
        {
            FileFunctions.WriteLine(_logFile, "---------------------------------------------------------------------------");
            FileFunctions.WriteLine(_logFile, message);
            FileFunctions.WriteLine(_logFile, "---------------------------------------------------------------------------");

            // Indicate search has finished.
            _dockPane.SearchStatus = DockpaneMainViewModel.SearchStatuses.NotStarted;
            _dockPane.ProgressUpdate(null, -1, -1);

            string imageSource = string.Format("pack://application:,,,/DataSearches;component/Images/{0}32.png", image);

            // Notify user of completion.
            Notification notification = new()
            {
                Title = "Data Searches",
                Severity = Notification.SeverityLevel.High,
                Message = message,
                ImageSource = new BitmapImage(new Uri(imageSource)) as ImageSource
            };
            FrameworkApplication.AddNotification(notification);

            // Open the log file (if required).
            if (OpenLogFile || _searchErrors)
                Process.Start("notepad.exe", _logFile);
        }

        /// <summary>
        /// Adds or removes the buffer layer to the active map and all other map windows.
        /// </summary>
        /// <param name="mapWindows">The list of additional map windows to update.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        private async Task<string> AddBufferLayerToMapsAsync(List<Map> mapWindows)
        {
            FileFunctions.WriteLine(_logFile, "");

            string bufferSymbologyFile = _layerPath + "\\" + _bufferLayerFile;

            // Add the buffer layer to all maps.
            string nameFromLyrx = await AddLayerToAllMapsAsync(_bufferLayerName, _bufferOutputFile, bufferSymbologyFile, mapWindows);
            if (nameFromLyrx == null)
                return null;

            // Turn off labels for the buffer layer in all maps.
            await TurnOffLabelsInAllMapsAsync(nameFromLyrx, mapWindows);

            FileFunctions.WriteLine(_logFile, "Buffer layer added to maps.");

            return nameFromLyrx;
        }

        /// <summary>
        /// Adds or removes the search area layer to the active map and all other map windows.
        /// </summary>
        /// <param name="mapWindows">The list of additional map windows to update.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        private async Task<string> AddSearchLayerToMapsAsync(List<Map> mapWindows)
        {
            FileFunctions.WriteLine(_logFile, "");

            string searchLayerFile = _searchSymbologyBase + _searchLayerExtension + ".lyrx";
            string searchSymbologyFile = _layerPath + "\\" + searchLayerFile;

            // Add the search feature layer to all maps.
            string nameFromLyrx = await AddLayerToAllMapsAsync(_searchLayerName, _searchOutputFile, searchSymbologyFile, mapWindows);
            if (nameFromLyrx == null)
                return null;

            // Turn off labels for the search feature layer in all maps.
            await TurnOffLabelsInAllMapsAsync(nameFromLyrx, mapWindows);

            FileFunctions.WriteLine(_logFile, "Search feature layer added to maps.");

            return nameFromLyrx;
        }

        /// <summary>
        /// Adds a layer with symbology to the active map and all other maps.
        /// </summary>
        private async Task<string> AddLayerToAllMapsAsync(string layerName, string outputFile, string symbologyFile, List<Map> mapWindows)
        {
            string nameFromLyrx;

            // Apply any symbology to the layer and move it to the group layer.
            nameFromLyrx = await SetLayerInMapAsync(layerName, symbologyFile, 0, null);
            if (nameFromLyrx == null)
            {
                FileFunctions.WriteLine(_logFile, $"Error setting layer '{layerName}' in the active map.");
                _searchErrors = true;

                return null;
            }

            // Add to all other map windows.
            foreach (Map map in mapWindows)
            {
                // Add the layer to the map.
                if (!await _mapFunctions.AddLayerToMapAsync(outputFile, 0, layerName, map))
                {
                    FileFunctions.WriteLine(_logFile, $"Error adding layer '{layerName}' to map: {map.Name}");
                    _searchErrors = true;
                }

                // Apply any symbology to the layer and move it to the group layer.
                if (await SetLayerInMapAsync(layerName, symbologyFile, 0, map) == null)
                {
                    FileFunctions.WriteLine(_logFile, $"Error setting layer '{layerName}' in map: {map.Name}");
                    _searchErrors = true;

                    return null;
                }
            }

            // Return the name of the layer from the symbology file.
            return nameFromLyrx;
        }

        /// <summary>
        /// Turns off labels for the specified layer in the active map and all other maps.
        /// </summary>
        private async Task TurnOffLabelsInAllMapsAsync(string layerName, List<Map> mapWindows)
        {
            if (!await _mapFunctions.SwitchLabelsAsync(layerName, false))
            {
                FileFunctions.WriteLine(_logFile, $"Error turning off labels for layer '{layerName}' in the active map.");
                _searchErrors = true;
            }

            foreach (Map map in mapWindows)
            {
                if (!await _mapFunctions.SwitchLabelsAsync(layerName, false, map))
                {
                    FileFunctions.WriteLine(_logFile, $"Error turning off labels for layer '{layerName}' in map: {map.Name}");
                    _searchErrors = true;
                }
            }
        }

        /// <summary>
        /// Removes the layer from the active map and all other maps, and deletes the associated feature class if provided.
        /// </summary>
        private async Task RemoveLayerFromAllMapsAsync(string layerName, string deletePath, List<Map> mapWindows)
        {
            try
            {
                // Remove from active map.
                await _mapFunctions.RemoveLayerAsync(layerName, null);

                // Remove from other map windows.
                foreach (Map map in mapWindows)
                {
                    await _mapFunctions.RemoveLayerAsync(layerName, map);
                }

                // Optionally delete the source data.
                if (!string.IsNullOrWhiteSpace(deletePath))
                {
                    await ArcGISFunctions.DeleteFeatureClassAsync(deletePath);
                }

                FileFunctions.WriteLine(_logFile, $"{layerName} layer deleted.");
            }
            catch
            {
                FileFunctions.WriteLine(_logFile, $"Error removing the {layerName} layer.");
                _searchErrors = true;
            }
        }

        /// <summary>
        /// Clean up after the search has completed (successfully or not).
        /// </summary>
        /// <returns></returns>
        private async Task CleanUpSearchAsync()
        {
            // Remove all temporary feature classes and tables from the active map.
            await _mapFunctions.RemoveLayerAsync(_tempMasterLayerName, null);
            await _mapFunctions.RemoveLayerAsync(_tempFCLayerName, null);
            await _mapFunctions.RemoveLayerAsync(_tempFCPointsLayerName, null);
            await _mapFunctions.RemoveLayerAsync(_tempSearchPointsLayerName, null);
            await _mapFunctions.RemoveTableAsync(_tempTableLayerName, null);

            // Delete the temporary feature classes and tables.
            if (await ArcGISFunctions.FeatureClassExistsAsync(_tempMasterOutputFile))
                await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempMasterLayerName);

            if (await ArcGISFunctions.FeatureClassExistsAsync(_tempFCOutputFile))
                await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempFCLayerName);

            if (await ArcGISFunctions.FeatureClassExistsAsync(_tempFCPointsOutputFile))
                await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempFCPointsLayerName);

            if (await ArcGISFunctions.FeatureClassExistsAsync(_tempSearchPointsOutputFile))
                await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempSearchPointsLayerName);

            if (await ArcGISFunctions.TableExistsAsync(_tempTableOutputFile))
                await ArcGISFunctions.DeleteGeodatabaseTableAsync(_tempGDBName, _tempTableLayerName);

            // Clear the search features selection in the active map.
            await _mapFunctions.ClearLayerSelectionAsync(_inputLayerName, null);

            // Remove the group layer from the active map if it is empty.
            await _mapFunctions.RemoveGroupLayerAsync(_groupLayerName, null);
        }

        /// <summary>
        /// Create the output folders if required.
        /// </summary>
        /// <param name="saveRootDir"></param>
        /// <param name="saveFolder"></param>
        /// <param name="extractFolder"></param>
        /// <param name="gisFolder"></param>
        /// <returns></returns>
        private bool CreateOutputFolders(string saveRootDir, string saveFolder, string extractFolder, string gisFolder, ref string extractPath, ref string gisPath)
        {
            extractPath = null;
            gisPath = null;

            // Create root folder if required.
            if (!FileFunctions.DirExists(saveRootDir))
            {
                try
                {
                    Directory.CreateDirectory(saveRootDir);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Cannot create directory '" + saveRootDir + "'. System error: " + ex.Message);
                    return false;
                }
            }

            // Create save sub-folder if required.
            if (!string.IsNullOrEmpty(saveFolder))
                saveFolder = saveRootDir + @"\" + saveFolder;
            else
                saveFolder = saveRootDir;
            if (!FileFunctions.DirExists(saveFolder))
            {
                try
                {
                    Directory.CreateDirectory(saveFolder);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Cannot create directory '" + saveFolder + "'. System error: " + ex.Message);
                    return false;
                }
            }

            // Create extract sub-folder if required.
            if (!string.IsNullOrEmpty(extractFolder))
                extractFolder = saveFolder + @"\" + extractFolder;
            else
                extractFolder = saveFolder;

            if (!FileFunctions.DirExists(extractFolder))
            {
                try
                {
                    Directory.CreateDirectory(extractFolder);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Cannot create directory '" + extractFolder + "'. System error: " + ex.Message);
                    return false;
                }
            }

            // Create gis sub-folder if required.
            if (!string.IsNullOrEmpty(gisFolder))
                gisFolder = saveFolder + @"\" + gisFolder;
            else
                gisFolder = saveFolder;

            if (!FileFunctions.DirExists(gisFolder))
            {
                try
                {
                    Directory.CreateDirectory(gisFolder);
                }
                catch (Exception ex)
                {
                    FileFunctions.WriteLine(_logFile, "Cannot create directory '" + gisFolder + "'. System error: " + ex.Message);
                    return false;
                }
            }

            extractPath = extractFolder;
            gisPath = gisFolder;

            return true;
        }

        /// <summary>
        /// Create an empty combined sites table (with just the column headers).
        /// </summary>
        /// <param name="combinedSitesTable"></param>
        /// <param name="combinedSitesTableOption"></param>
        /// <returns></returns>
        private bool CreateCombinedSitesTable(string combinedSitesTable, CombinedSitesTableOptions combinedSitesTableOption)
        {
            // Start the table if overwrite has been selected, or if the table doesn't exist (and append has been selected).
            if (combinedSitesTableOption == CombinedSitesTableOptions.Overwrite ||
                combinedSitesTableOption == CombinedSitesTableOptions.Append && !FileFunctions.FileExists(combinedSitesTable))
            {
                if (!FileFunctions.WriteEmptyTextFile(combinedSitesTable, _combinedSitesColumnList))
                {
                    FileFunctions.WriteLine(_logFile, "Error writing to combined sites table");
                    _searchErrors = true;

                    return false;
                }

                FileFunctions.WriteLine(_logFile, "Combined sites table started");
            }

            return true;
        }

        /// <summary>
        /// Find the search reference features in the search layers.
        /// </summary>
        /// <param name="reference"></param>
        /// <returns>Name of the target layer.</returns>
        private async Task<bool> FindSearchFeaturesAsync(string searchClause)
        {
            // Loop through all base layer and extension combinations.
            foreach (string searchLayerExtension in _searchLayerExtensions)
            {
                string searchLayer = _searchLayerBase + searchLayerExtension;

                // Find the feature layer by name in the search map if it exists. Only search existing layers.
                FeatureLayer featureLayer = await _mapFunctions.FindLayerAsync(searchLayer, null);

                if (featureLayer != null)
                {
                    // Count the required features in the layer.
                    long featureCount = await ArcGISFunctions.GetFeaturesCountAsync(featureLayer, searchClause);

                    if (featureCount > 0)
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Count the search reference features in the search layers.
        /// </summary>
        /// <param name="reference"></param>
        /// <returns>Name of the target layer.</returns>
        private async Task<long> CountSearchFeaturesAsync(string searchClause)
        {
            // Find the search feature.
            int featureLayerCount = 0;
            long totalFeatureCount = 0;

            // Loop through all base layer and extension combinations.
            foreach (string searchLayerExtension in _searchLayerExtensions)
            {
                string searchLayer = _searchLayerBase + searchLayerExtension;

                // Find the feature layer by name in the active map if it exists. Only search existing layers.
                FeatureLayer featureLayer = await _mapFunctions.FindLayerAsync(searchLayer, null);

                if (featureLayer != null)
                {
                    // Count the required features in the layer.
                    long featureCount = await ArcGISFunctions.GetFeaturesCountAsync(featureLayer, searchClause);

                    if (featureCount == 0)
                        FileFunctions.WriteLine(_logFile, "No features found in " + searchLayer);

                    if (featureCount > 0)
                    {
                        FileFunctions.WriteLine(_logFile, featureCount.ToString() + " feature(s) found in " + searchLayer);
                        if (featureLayerCount == 0)
                        {
                            // Save the layer name and extension where the feature(s) were found.
                            _inputLayerName = searchLayer;
                            _searchLayerExtension = searchLayerExtension;
                        }
                        totalFeatureCount += featureCount;
                        featureLayerCount++;
                    }
                }
            }

            // If no features found in any layer.
            if (featureLayerCount == 0)
            {
                FileFunctions.WriteLine(_logFile, "No features found in any of the search layers");

                return 0;
            }

            // If features found in more than one layer.
            if (featureLayerCount > 1)
            {
                FileFunctions.WriteLine(_logFile, totalFeatureCount.ToString() + " features found in different search layers");

                return 0;
            }

            // If multiple features found.
            if (totalFeatureCount > 1)
            {
                // Ask the user if they want to continue
                MessageBoxResult response = MessageBox.Show(totalFeatureCount.ToString() + " features found in " + _inputLayerName + " matching those criteria. Do you wish to continue?", _displayName, MessageBoxButton.YesNo);
                if (response == MessageBoxResult.No)
                {
                    FileFunctions.WriteLine(_logFile, totalFeatureCount.ToString() + " features found in the search layers");

                    return 0;
                }
            }

            return totalFeatureCount;
        }

        /// <summary>
        /// Prepare a new temporary GDB to use and check it's empty (in case it
        /// already existed).
        /// </summary>
        /// <returns></returns>
        private async Task<bool> PrepareTemporaryGDBAsync()
        {
            // Set a temporary folder path.
            string tempFolder = Path.GetTempPath();

            // Create the temporary file geodatabase if it doesn't exist.
            _tempGDBName = tempFolder + @"Temp.gdb";

            Geodatabase tempGDB;
            bool tempGDBFound = true;
            if (!FileFunctions.DirExists(_tempGDBName))
            {
                tempGDB = ArcGISFunctions.CreateFileGeodatabase(_tempGDBName);
                if (tempGDB == null)
                {
                    FileFunctions.WriteLine(_logFile, "Error creating temporary geodatabase " + _tempGDBName);
                    _searchErrors = true;

                    return false;
                }

                tempGDBFound = false;
                FileFunctions.WriteLine(_logFile, "Temporary geodatabase created");
            }

            // Set the temporary layer and file names.
            _tempMasterLayerName = "TempMaster_" + _userID;
            _tempMasterOutputFile = _tempGDBName + @"\" + _tempMasterLayerName;
            _tempFCLayerName = "TempOutput_" + _userID;
            _tempFCOutputFile = _tempGDBName + @"\" + _tempFCLayerName;
            _tempFCPointsLayerName = "TempOutputPoints_" + _userID;
            _tempFCPointsOutputFile = _tempGDBName + @"\" + _tempFCPointsLayerName;
            _tempSearchPointsLayerName = "TempSearchPoints_" + _userID;
            _tempSearchPointsOutputFile = _tempGDBName + @"\" + _tempSearchPointsLayerName;
            _tempTableLayerName = "TempTable_" + _userID;
            _tempTableOutputFile = _tempGDBName + @"\" + _tempTableLayerName;

            // If the GDB already existed clean it up.
            if (tempGDBFound)
            {
                // Delete the temporary master feature class from the active map if it still exists.
                await _mapFunctions.RemoveLayerAsync(_tempMasterLayerName, null);
                if (await ArcGISFunctions.FeatureClassExistsAsync(_tempMasterOutputFile))
                {
                    await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempMasterLayerName);
                    //FileFunctions.WriteLine(_logFile, "Temporary master feature class deleted");
                }

                // Delete the temporary output feature class from the active map if it still exists.
                await _mapFunctions.RemoveLayerAsync(_tempFCLayerName, null);
                if (await ArcGISFunctions.FeatureClassExistsAsync(_tempFCOutputFile))
                {
                    await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempFCLayerName);
                    //FileFunctions.WriteLine(_logFile, "Temporary output feature class deleted");
                }

                // Delete the temporary output points feature class from the active map if it still exists.
                await _mapFunctions.RemoveLayerAsync(_tempFCPointsLayerName, null);
                if (await ArcGISFunctions.FeatureClassExistsAsync(_tempFCPointsOutputFile))
                {
                    await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempFCPointsLayerName);
                    //FileFunctions.WriteLine(_logFile, "Temporary output feature class deleted");
                }

                // Delete the temporary search points feature class from the active map if it still exists.
                await _mapFunctions.RemoveLayerAsync(_tempSearchPointsLayerName, null);
                if (await ArcGISFunctions.FeatureClassExistsAsync(_tempSearchPointsOutputFile))
                {
                    await ArcGISFunctions.DeleteGeodatabaseFCAsync(_tempGDBName, _tempSearchPointsLayerName);
                    //FileFunctions.WriteLine(_logFile, "Temporary output feature class deleted");
                }

                // Delete the temporary output table from the active map if it still exists.
                await _mapFunctions.RemoveTableAsync(_tempTableLayerName, null);
                if (await ArcGISFunctions.TableExistsAsync(_tempTableOutputFile))
                {
                    await ArcGISFunctions.DeleteGeodatabaseTableAsync(_tempGDBName, _tempTableLayerName);
                    //FileFunctions.WriteLine(_logFile, "Temporary output table deleted");
                }
            }

            return true;
        }

        /// <summary>
        /// Save the selected search feature(s) to a new layer.
        /// </summary>
        /// <returns></returns>
        private async Task<bool> SaveSearchFeaturesAsync()
        {
            // Get the full layer path (in case it's nested in one or more groups).
            //string inputLayerPath = _mapFunctions.GetLayerPath(_inputLayerName);

            // Move potentially serialization-prone code to the CIM thread.
            string inputLayerPath = await QueuedTask.Run(async () =>
            {
                return await _mapFunctions.GetLayerPathAsync(_inputLayerName);
            });

            FileFunctions.WriteLine(_logFile, "Saving search feature(s)");

            // Copy the selected feature(s) to an output file.
            if (!await ArcGISFunctions.CopyFeaturesAsync(inputLayerPath, _searchOutputFile, true))
            {
                FileFunctions.WriteLine(_logFile, "Error saving search feature(s)");
                _searchErrors = true;

                return false;
            }

            return true;
        }

        /// <summary>
        /// Buffer the selected search feature(s) to a new layer.
        /// </summary>
        /// <param name="bufferSize"></param>
        /// <param name="bufferUnitProcess"></param>
        /// <param name="bufferUnitShort"></param>
        /// <returns></returns>
        private async Task<bool> BufferSearchFeaturesAsync(string bufferSize, string bufferUnitProcess, string bufferUnitShort)
        {
            // Get the full layer path (in case it's nested in one or more groups).
            //string searchLayerPath = _mapFunctions.GetLayerPath(_searchLayerName);

            // Move potentially serialization-prone code to the CIM thread.
            string searchLayerPath = await QueuedTask.Run(async () =>
            {
                return await _mapFunctions.GetLayerPathAsync(_searchLayerName);
            });

            // Create a buffer around the feature and save into a new file.
            string bufferDistance = bufferSize + " " + bufferUnitProcess;

            // Safeguard for zero buffer size; Select a tiny buffer to allow
            // correct legending (expects a polygon).
            if (bufferSize == "0")
                bufferDistance = "0.01 Meters";

            // Check if all fields in the aggregate fields exist. If not, ignore.
            List<string> aggColumns = [.. _bufferFields.Split(';')];
            string dissolveFields = "";
            foreach (string fieldName in aggColumns)
            {
                if (await _mapFunctions.FieldExistsAsync(searchLayerPath, fieldName, null))
                    dissolveFields = dissolveFields + fieldName + ";";
            }

            FileFunctions.WriteLine(_logFile, "Buffering feature(s) with a distance of " + bufferSize + bufferUnitShort);

            string dissolveOption = "ALL";
            if (!string.IsNullOrEmpty(dissolveFields))
            {
                dissolveFields = dissolveFields.Substring(0, dissolveFields.Length - 1);
                dissolveOption = "LIST";
            }

            if (!await ArcGISFunctions.BufferFeaturesAsync(searchLayerPath, _bufferOutputFile, bufferDistance, "FULL", "ROUND", dissolveOption, dissolveFields, addToMap: true))
            {
                FileFunctions.WriteLine(_logFile, "Error during feature buffering");
                _searchErrors = true;

                return false;
            }

            return true;
        }

        /// <summary>
        /// Apply any symbology to the new layer and move it to the group layer.
        /// </summary>
        /// <param name="layerName"></param>
        /// <param name="symbologyFile"></param>
        /// <param name="position"></param>
        /// <returns></returns>
        private async Task<string> SetLayerInMapAsync(string layerName, string symbologyFile, int position = -1, Map targetMap = null)
        {
            string nameFromLyrx = layerName;

            // Apply layer symbology.
            if (!string.IsNullOrEmpty(symbologyFile))
            {
                if (FileFunctions.FileExists(symbologyFile) && symbologyFile.Substring(symbologyFile.Length - 4, 4).Equals("lyrx", StringComparison.OrdinalIgnoreCase))
                {
                    //FileFunctions.WriteLine(_logFile, "Applying symbology to '" + layerName + "'");

                    nameFromLyrx = await _mapFunctions.ApplySymbologyFromLayerFileAsync(layerName, symbologyFile, targetMap);
                    if (string.IsNullOrEmpty(nameFromLyrx))
                    {
                        FileFunctions.WriteLine(_logFile, "Error applying symbology to '" + layerName + "'");
                        _searchErrors = true;

                        return null;
                    }
                }
                else
                {
                    FileFunctions.WriteLine(_logFile, "Error layer file '" + symbologyFile + "' not found or valid");
                }
            }

            // Move layer to the group layer.
            if (!string.IsNullOrEmpty(_groupLayerName))
            {
                if (!await _mapFunctions.MoveToGroupLayerAsync(await _mapFunctions.FindLayerAsync(nameFromLyrx, targetMap), _groupLayerName, position, targetMap))
                {
                    FileFunctions.WriteLine(_logFile, "Error moving layer to '" + _groupLayerName + "'");
                    _searchErrors = true;

                    return null;
                }
            }

            // Return the name of the layer in case it was changed in the symbology file.
            return nameFromLyrx;
        }

        /// <summary>
        /// Process each of the selected map layers.
        /// </summary>
        /// <param name="selectedLayer"></param>
        /// <param name="reference"></param>
        /// <param name="siteName"></param>
        /// <param name="shortRef"></param>
        /// <param name="subref"></param>
        /// <param name="radius"></param>
        /// <param name="areaMeasureUnit"></param>
        /// <param name="keepSelectedLayers"></param>
        /// <param name="addSelectedLayersOption"></param>
        /// <param name="overwriteLabelOption"></param>
        /// <param name="combinedSitesTableOption"></param>
        /// <returns></returns>
        private async Task<bool> ProcessMapLayerAsync(MapLayer selectedLayer, string reference, string siteName,
            string shortRef, string subref, string radius, string organisation, string areaMeasureUnit,
            bool keepSelectedLayers,
            AddSelectedLayersOptions addSelectedLayersOption,
            OverwriteLabelOptions overwriteLabelOption,
            CombinedSitesTableOptions combinedSitesTableOption)
        {
            // Get the settings relevant for this layer.
            string mapNodeGroup = selectedLayer.NodeGroup;
            //string mapNodeLayer = selectedLayer.NodeLayer;
            string mapLayerName = selectedLayer.LayerName;
            string mapMapName = selectedLayer.MapName;
            string mapOutputName = selectedLayer.GISOutputName;
            string mapTableOutputName = selectedLayer.TableOutputName;
            string mapColumns = selectedLayer.Columns;
            string mapGroupColumns = selectedLayer.GroupColumns;
            string mapStatsColumns = selectedLayer.StatisticsColumns;
            string mapOrderColumns = selectedLayer.OrderColumns;
            string mapCriteria = selectedLayer.Criteria;

            bool mapIncludeArea = selectedLayer.IncludeArea;
            string mapIncludeNearFields = selectedLayer.IncludeNearFields;
            bool mapIncludeRadius = selectedLayer.IncludeRadius;

            string mapKeyColumn = selectedLayer.KeyColumn;
            string mapFormat = selectedLayer.Format;
            bool mapKeepLayer = selectedLayer.KeepLayer;
            string mapOutputType = selectedLayer.OutputType;

            bool mapDisplayLabels = selectedLayer.DisplayLabels;
            string mapLayerFileName = selectedLayer.LayerFileName;
            bool mapOverwriteLabels = selectedLayer.OverwriteLabels;
            string mapLabelColumn = selectedLayer.LabelColumn;
            string mapLabelClause = selectedLayer.LabelClause;
            string mapMacroName = selectedLayer.MacroName;

            string mapCombinedSitesColumns = selectedLayer.CombinedSitesColumns;
            string mapCombinedSitesGroupColumns = selectedLayer.CombinedSitesGroupColumns;
            string mapCombinedSitesStatsColumns = selectedLayer.CombinedSitesStatisticsColumns;
            string mapCombinedSitesOrderColumns = selectedLayer.CombinedSitesOrderByColumns;

            // Deal with wildcards in the output names.
            mapOutputName = StringFunctions.ReplaceSearchStrings(mapOutputName, reference, siteName, shortRef, subref, radius, organisation);
            mapTableOutputName = StringFunctions.ReplaceSearchStrings(mapTableOutputName, reference, siteName, shortRef, subref, radius, organisation);

            // Remove any illegal characters from the names.
            mapOutputName = StringFunctions.StripIllegals(mapOutputName, _repChar);
            mapTableOutputName = StringFunctions.StripIllegals(mapTableOutputName, _repChar);

            // Set the statistics columns if they haven't been supplied.
            if (string.IsNullOrEmpty(mapStatsColumns) && !string.IsNullOrEmpty(mapGroupColumns))
                mapStatsColumns = StringFunctions.AlignStatsColumns(mapColumns, mapStatsColumns, mapGroupColumns);
            if (string.IsNullOrEmpty(mapCombinedSitesStatsColumns) && !string.IsNullOrEmpty(mapCombinedSitesGroupColumns))
                mapCombinedSitesStatsColumns = StringFunctions.AlignStatsColumns(mapCombinedSitesColumns, mapCombinedSitesStatsColumns, mapCombinedSitesGroupColumns);

            // Get the full layer path (in case it's nested in one or more groups).
            string mapLayerPath = await _mapFunctions.GetLayerPathAsync(mapLayerName);

            // Find the map layer by name in the active map.
            FeatureLayer mapLayer = await _mapFunctions.FindLayerAsync(mapLayerName, null);
            if (mapLayer == null)
            {
                FileFunctions.WriteLine(_logFile, "Error finding map layer '" + mapLayerName + "' in the active map");
                _searchErrors = true;

                return false;
            }

            // Get the target map for the map name.
            Map targetMap = null;
            if (!string.IsNullOrWhiteSpace(mapMapName))
            {
                targetMap = await _mapFunctions.GetMapFromNameAsync(mapMapName);
                if (targetMap == null)
                {
                    FileFunctions.WriteLine(_logFile, "Error finding map '" + mapMapName + "'");
                    _searchErrors = true;

                    return false;
                }
            }

            // Select by location.
            FileFunctions.WriteLine(_logFile, "Selecting features using selected feature(s) from layer '" + _bufferLayerName + "' ...");
            if (!await MapFunctions.SelectLayerByLocationAsync(mapLayer, _bufferLayer, "INTERSECT", "", "NEW_SELECTION"))
            {
                FileFunctions.WriteLine(_logFile, "Error selecting layer '" + mapLayerName + "' by location");
                _searchErrors = true;

                return false;
            }

            // Refine the selection by attributes (if required).
            int selectionCount = await QueuedTask.Run(() => mapLayer.SelectionCount);
            if (selectionCount > 0 && !string.IsNullOrEmpty(mapCriteria))
            {
                FileFunctions.WriteLine(_logFile, "Refining selection with criteria " + mapCriteria + " ...");

                if (!await _mapFunctions.SelectLayerByAttributesAsync(mapLayerName, mapCriteria, SelectionCombinationMethod.And))
                {
                    FileFunctions.WriteLine(_logFile, "Error refining selection on layer " + mapLayerName + " with criteria " + mapCriteria + ". Please check syntax and column names (case sensitive)");
                    _searchErrors = true;

                    return false;
                }
            }

            // Count the selected features.
            int featureCount = await QueuedTask.Run(() => mapLayer.SelectionCount);

            // Write out the results - to a feature class initially. Include distance if required.
            if (featureCount <= 0)
            {
                FileFunctions.WriteLine(_logFile, "No features found");
                return true;
            }

            FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", featureCount) + " feature(s) found");

            // Create the temporary map output depending on the output type required.
            if (!await CreateMapOutputAsync(mapLayerName, mapLayerPath, _bufferLayerPath, mapOutputType))
            {
                MessageBox.Show("Cannot output selection from " + mapLayerName + " to " + _tempMasterOutputFile + ".");
                FileFunctions.WriteLine(_logFile, "Cannot output selection from " + mapLayerName + " to " + _tempMasterOutputFile);

                return false;
            }

            // Add map labels to the temporary output if required and being kept.
            if (!string.IsNullOrEmpty(mapLabelColumn) && (keepSelectedLayers) && (mapKeepLayer))
            {
                if (!await AddMapLabelsAsync(overwriteLabelOption, mapOverwriteLabels, mapLabelColumn, mapKeyColumn, mapNodeGroup))
                {
                    FileFunctions.WriteLine(_logFile, "Error adding map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }
            }

            // Create relevant output names.
            string mapOutputFile = _gisPath + @"\" + mapOutputName; // Output shapefile/feature class name. Note no extension to allow write to GDB.
            string mapTableOutputFile = _outputPath + @"\" + mapTableOutputName + "." + mapFormat.ToLower(); // Output table name.

            // Include headers for CSV files.
            bool includeHeaders = false;
            if (mapFormat.Equals("csv", StringComparison.OrdinalIgnoreCase))
                includeHeaders = true;

            // Only include radius if requested.
            string radiusText = "none";
            if (mapIncludeRadius)
                radiusText = radius;

            string areaUnit = "";
            if (mapIncludeArea)
                areaUnit = areaMeasureUnit;

            // Export results to table if required.
            if (!string.IsNullOrEmpty(mapFormat) && !string.IsNullOrEmpty(mapColumns))
            {
                FileFunctions.WriteLine(_logFile, "Extracting summary information ...");

                int intLineCount = await ExportSelectionAsync(mapTableOutputFile, mapFormat.ToLower(), mapColumns, mapGroupColumns, mapStatsColumns, mapOrderColumns,
                    includeHeaders, false, areaUnit, mapIncludeNearFields, radiusText);
                if (intLineCount <= 0)
                {
                    FileFunctions.WriteLine(_logFile, "Error extracting summary from " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }

                FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " record(s) exported");
            }

            // If selected layers are to be kept, and this layer is to be kept,
            // copy to a permanent layer and add it to the target map (if required).
            if ((keepSelectedLayers) && (mapKeepLayer))
            {
                if (!await KeepLayerAsync(mapOutputName, mapOutputFile, addSelectedLayersOption, mapLayerFileName, mapDisplayLabels, mapLabelClause,
                    mapLabelColumn, targetMap))
                {
                    _searchErrors = true;

                    return false;
                }
            }

            // Add to combined sites table if required.
            if (!string.IsNullOrEmpty(mapCombinedSitesColumns) && combinedSitesTableOption != CombinedSitesTableOptions.None)
            {
                FileFunctions.WriteLine(_logFile, "Extracting summary output for combined sites table ...");

                int intLineCount = await ExportSelectionAsync(_combinedSitesOutputFile, _combinedSitesTableFormat, mapCombinedSitesColumns, mapCombinedSitesGroupColumns,
                    mapCombinedSitesStatsColumns, mapCombinedSitesOrderColumns,
                    false, true, areaUnit, mapIncludeNearFields, radiusText);

                if (intLineCount < 0)
                {
                    FileFunctions.WriteLine(_logFile, "Error extracting summary for combined sites table from " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }

                FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " row(s) added to combined sites table");
            }

            // Clear the selection in the input layer in the active map.
            await _mapFunctions.ClearLayerSelectionAsync(mapLayerName, null);

            FileFunctions.WriteLine(_logFile, "Analysis complete");

            // Trigger the macro if one exists
            if (!string.IsNullOrEmpty(mapMacroName))
            {
                FileFunctions.WriteLine(_logFile, "Executing vbscript macro ...");

                if (!StartProcess(mapMacroName, mapTableOutputName, mapFormat))
                {
                    FileFunctions.WriteLine(_logFile, "Error executing vbscript macro '" + mapMacroName + "'.");
                    _searchErrors = true;

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Create the required temporary output type from the current layer.
        /// </summary>
        /// <param name="mapLayerName"></param>
        /// <param name="mapLayerPath"></param>
        /// <param name="bufferLayerPath"></param>
        /// <param name="mapOutputType"></param>
        /// <returns></returns>
        private async Task<bool> CreateMapOutputAsync(string mapLayerName, string mapLayerPath, string bufferLayerPath, string mapOutputType)
        {
            // Get the input feature class type.
            string mapLayerFCType = await _mapFunctions.GetFeatureClassTypeAsync(mapLayerName);
            if (mapLayerFCType == null)
                return false;

            // Get the buffer feature class type.
            string bufferFCType = await _mapFunctions.GetFeatureClassTypeAsync(_bufferLayerName);
            if (bufferFCType == null)
                return false;

            // If the input layer should be clipped to the buffer layer, do so now.
            if (mapOutputType.Equals("CLIP", StringComparison.OrdinalIgnoreCase))
            {
                if (mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) && bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) ||
                    mapLayerFCType.Equals("line", StringComparison.OrdinalIgnoreCase) &&
                    (bufferFCType.Equals("line", StringComparison.OrdinalIgnoreCase) || bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase)))
                {
                    // Clip
                    FileFunctions.WriteLine(_logFile, "Clipping selected features ...");
                    return await ArcGISFunctions.ClipFeaturesAsync(mapLayerPath, bufferLayerPath, _tempMasterOutputFile, true);
                }
                else
                {
                    // Copy
                    FileFunctions.WriteLine(_logFile, "Copying selected features ...");
                    return await ArcGISFunctions.CopyFeaturesAsync(mapLayerPath, _tempMasterOutputFile, true);
                }
            }
            // If the buffer layer should be clipped to the input layer, do so now.
            else if (mapOutputType.Equals("OVERLAY", StringComparison.OrdinalIgnoreCase))
            {
                if (bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) && mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) ||
                    bufferFCType.Equals("line", StringComparison.OrdinalIgnoreCase) &&
                    (mapLayerFCType.Equals("line", StringComparison.OrdinalIgnoreCase) || mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase)))
                {
                    // Clip
                    FileFunctions.WriteLine(_logFile, "Overlaying selected features ...");
                    return await ArcGISFunctions.ClipFeaturesAsync(bufferLayerPath, mapLayerPath, _tempMasterOutputFile, true);
                }
                else
                {
                    // Find the map layer by name in the target map.
                    FeatureLayer mapLayer = await _mapFunctions.FindLayerAsync(mapLayerName, null);
                    if (mapLayer == null)
                    {
                        FileFunctions.WriteLine(_logFile, "Error finding map layer '" + mapLayerName + "' in map '" + mapLayerName + "'");
                        _searchErrors = true;

                        return false;
                    }

                    // Find the buffer layer by name in the target map.
                    FeatureLayer bufferLayer = await _mapFunctions.FindLayerAsync(_bufferLayerName, null);
                    if (bufferLayer == null)
                    {
                        FileFunctions.WriteLine(_logFile, "Error finding buffer layer '" + _bufferLayerName + "' in active map");
                        _searchErrors = true;

                        return false;
                    }

                    // Select from the buffer layer.
                    FileFunctions.WriteLine(_logFile, "Selecting features  ...");
                    await MapFunctions.SelectLayerByLocationAsync(bufferLayer, mapLayer);

                    if (bufferLayer == null)
                        return false;

                    // Count the selected features.
                    int featureCount = await QueuedTask.Run(() => bufferLayer.SelectionCount);
                    if (featureCount > 0)
                    {
                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", featureCount) + " feature(s) found");

                        // Copy the selection from the buffer layer.
                        FileFunctions.WriteLine(_logFile, "Copying selected features ... ");
                        if (!await ArcGISFunctions.CopyFeaturesAsync(bufferLayerPath, _tempMasterOutputFile, true))
                            return false;
                    }
                    else
                    {
                        FileFunctions.WriteLine(_logFile, "No features selected");

                        return true;
                    }

                    // Clear the buffer layer selection in the active map.
                    await _mapFunctions.ClearLayerSelectionAsync(_bufferLayerName, null);

                    return true;
                }
            }
            // If the input layer should be intersected with the buffer layer, do so now.
            else if (mapOutputType.Equals("INTERSECT", StringComparison.OrdinalIgnoreCase))
            {
                if (mapLayerFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) && bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase) ||
                    mapLayerFCType.Equals("line", StringComparison.OrdinalIgnoreCase) &&
                    (bufferFCType.Equals("line", StringComparison.OrdinalIgnoreCase) || bufferFCType.Equals("polygon", StringComparison.OrdinalIgnoreCase)))
                {
                    string[] features = ["'" + mapLayerPath + "' #", "'" + bufferLayerPath + "' #"];
                    string inFeatures = string.Join(";", features);

                    // Intersect
                    FileFunctions.WriteLine(_logFile, "Intersecting selected features ...");
                    return await ArcGISFunctions.IntersectFeaturesAsync(inFeatures, _tempMasterOutputFile, addToMap: true); // Selected features in input, buffer FC, output.
                }
                else
                {
                    // Copy
                    FileFunctions.WriteLine(_logFile, "Copying selected features ...");
                    return await ArcGISFunctions.CopyFeaturesAsync(mapLayerPath, _tempMasterOutputFile, true);
                }
            }
            // Otherwise do a straight copy of the input layer.
            else
            {
                // Copy
                FileFunctions.WriteLine(_logFile, "Copying selected features ...");
                return await ArcGISFunctions.CopyFeaturesAsync(mapLayerPath, _tempMasterOutputFile, true);
            }
        }

        /// <summary>
        /// Add the map labels to the current layer.
        /// </summary>
        /// <param name="overwriteLabelOption"></param>
        /// <param name="overwriteLabels"></param>
        /// <param name="mapLabelColumn"></param>
        /// <param name="mapKeyColumn"></param>
        /// <param name="mapGroupName"></param>
        /// <returns></returns>
        private async Task<bool> AddMapLabelsAsync(OverwriteLabelOptions overwriteLabelOption, bool overwriteLabels,
            string mapLabelColumn, string mapKeyColumn, string mapGroupName)
        {
            bool newLabelField = false;
            // Does the map label field already exist? If not, add it.
            if (!await _mapFunctions.FieldExistsAsync(_tempMasterLayerName, mapLabelColumn, null))
            {
                if (!await ArcGISFunctions.AddFieldAsync(_tempMasterOutputFile, mapLabelColumn, "LONG"))
                {
                    FileFunctions.WriteLine(_logFile, "Error adding map label field '" + mapLabelColumn + "' to " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }

                newLabelField = true;
            }

            // Either we have a new label field, or we want to overwrite the labels and are allowed to.
            if (newLabelField ||
                overwriteLabelOption != OverwriteLabelOptions.No &&
                overwriteLabels)
            {
                // Add labels as required.
                if (!await CreateMapLabelsAsync(overwriteLabelOption, mapLabelColumn, mapKeyColumn, mapGroupName))
                {
                    FileFunctions.WriteLine(_logFile, "Error setting map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Create map labels for the current layer.
        /// </summary>
        /// <param name="overwriteLabelOption"></param>
        /// <param name="mapLabelColumn"></param>
        /// <param name="mapKeyColumn"></param>
        /// <param name="mapGroupName"></param>
        /// <returns></returns>
        private async Task<bool> CreateMapLabelsAsync(OverwriteLabelOptions overwriteLabelOption, string mapLabelColumn, string mapKeyColumn,
            string mapGroupName)
        {
            FileFunctions.WriteLine(_logFile, "Adding map labels ...");

            // Add relevant labels.
            if (overwriteLabelOption == OverwriteLabelOptions.ResetByLayer) // Reset each layer to 1.
            {
                FileFunctions.WriteLine(_logFile, "Resetting label counter ...");

                if (await _mapFunctions.AddIncrementalNumbersAsync(_tempMasterOutputFile, _tempMasterLayerName, mapLabelColumn, mapKeyColumn, 1) < 0)
                {
                    FileFunctions.WriteLine(_logFile, "Error calculating map label field '" + mapLabelColumn + "' in " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return false;
                }
            }
            else if (overwriteLabelOption == OverwriteLabelOptions.ResetByGroup && !string.IsNullOrEmpty(mapGroupName))
            {
                // Increment within but reset between groups. Note all group labels are already initialised as 1.
                // Only triggered if a group name has been found.
                int groupIndex = _mapGroupNames.IndexOf(mapGroupName);
                int groupLabel = _mapGroupLabels[groupIndex];

                groupLabel = await _mapFunctions.AddIncrementalNumbersAsync(_tempMasterOutputFile, _tempMasterLayerName, mapLabelColumn, mapKeyColumn, groupLabel);

                // Increment the new group label.
                groupLabel++;

                // Store the new group label.
                _mapGroupLabels[groupIndex] = groupLabel;
            }
            else
            {
                // There is no group or groups are ignored, or we are not resetting. Use the existing max label number.
                int startLabel = _maxLabel;

                _maxLabel = await _mapFunctions.AddIncrementalNumbersAsync(_tempMasterOutputFile, _tempMasterLayerName, mapLabelColumn, mapKeyColumn, startLabel);

                // Increment the new max label.
                _maxLabel++;
            }

            return true;
        }

        /// <summary>
        /// Export the selected features from the current layer to a text file.
        /// </summary>
        /// <param name="outputTableName"></param>
        /// <param name="outputFormat"></param>
        /// <param name="mapColumns"></param>
        /// <param name="mapGroupColumns"></param>
        /// <param name="mapStatsColumns"></param>
        /// <param name="mapOrderColumns"></param>
        /// <param name="includeHeaders"></param>
        /// <param name="append"></param>
        /// <param name="areaUnit"></param>
        /// <param name="includeNearFields"></param>
        /// <param name="radiusText"></param>
        /// <returns></returns>
        private async Task<int> ExportSelectionAsync(string outputTableName, string outputFormat,
            string mapColumns, string mapGroupColumns, string mapStatsColumns, string mapOrderColumns,
            bool includeHeaders, bool append, string areaUnit, string includeNearFields, string radiusText)
        {
            int intLineCount;

            // Only export if the user has specified columns.
            if (string.IsNullOrEmpty(mapColumns))
                return -1;

            // Check the input feature layer exists in the active map.
            FeatureLayer inputFeaturelayer = await _mapFunctions.FindLayerAsync(_tempMasterLayerName, null);
            if (inputFeaturelayer == null)
                return -1;

            // Get the input feature class type.
            string inputFeatureType = await _mapFunctions.GetFeatureClassTypeAsync(inputFeaturelayer);
            if (inputFeatureType == null)
                return -1;

            // Calculate the area field if required.
            string areaColumnName = "";
            if (!string.IsNullOrEmpty(areaUnit) && inputFeatureType.Equals("polygon", StringComparison.OrdinalIgnoreCase))
            {
                areaColumnName = "Area" + areaUnit;
                // Does the area field already exist? If not, add it.
                if (!await _mapFunctions.FieldExistsAsync(_tempMasterLayerName, areaColumnName, null))
                {
                    if (!await ArcGISFunctions.AddFieldAsync(_tempMasterOutputFile, areaColumnName, "DOUBLE", 20))
                    {
                        FileFunctions.WriteLine(_logFile, "Error adding area field to " + _tempMasterOutputFile);
                        _searchErrors = true;

                        return -1;
                    }

                    string geometryProperty = areaColumnName + " AREA";
                    if (areaUnit.Equals("ha", StringComparison.OrdinalIgnoreCase))
                    {
                        areaUnit = "HECTARES";
                    }
                    else if (areaUnit.Equals("m2", StringComparison.OrdinalIgnoreCase))
                    {
                        areaUnit = "SQUARE_METERS";
                    }
                    else if (areaUnit.Equals("km2", StringComparison.OrdinalIgnoreCase))
                    {
                        areaUnit = "SQUARE_KILOMETERS";
                    }

                    // Calculate the area field.
                    if (!await ArcGISFunctions.CalculateGeometryAsync(_tempMasterOutputFile, geometryProperty, "", areaUnit))
                    {
                        FileFunctions.WriteLine(_logFile, "Error calculating area field in " + _tempMasterOutputFile);
                        _searchErrors = true;

                        return -1;
                    }
                }
            }

            // Include radius if requested
            if (radiusText != "none")
            {
                FileFunctions.WriteLine(_logFile, "Including radius column ...");

                // Does the radius field already exist? If not, add it.
                if (!await _mapFunctions.FieldExistsAsync(_tempMasterLayerName, "Radius", null))
                {
                    if (!await ArcGISFunctions.AddFieldAsync(_tempMasterOutputFile, "Radius", "TEXT", fieldLength: 25))
                    {
                        FileFunctions.WriteLine(_logFile, "Error adding radius field to " + _tempMasterOutputFile);
                        _searchErrors = true;

                        return -1;
                    }
                }

                // Calculate the radius field.
                if (!await ArcGISFunctions.CalculateFieldAsync(_tempMasterOutputFile, "Radius", '"' + radiusText + '"'))
                {
                    FileFunctions.WriteLine(_logFile, "Error calculating radius field in " + _tempMasterOutputFile);
                    _searchErrors = true;

                    return -1;
                }
            }

            // Copy the input features.
            if (!await ArcGISFunctions.CopyFeaturesAsync(_tempMasterOutputFile, _tempFCOutputFile, true))
            {
                FileFunctions.WriteLine(_logFile, "Error copying output file to " + _tempFCOutputFile);
                _searchErrors = true;

                return -1;
            }

            //-------------------------------------------------------------
            // After this the input to the remainder of the function
            // should be reading from _tempFCOutputFile (_tempFCLayerName).
            //-------------------------------------------------------------

            // Calculate the boundary distance and bearing if required.
            if (includeNearFields.Equals("BOUNDARY", StringComparison.OrdinalIgnoreCase))
            {
                // Calculate the distance and additional proximity fields.
                if (!await ArcGISFunctions.NearAnalysisAsync(_tempFCOutputFile, _searchLayerName,
                    radiusText, "LOCATION", "ANGLE", "PLANAR", null, "METERS"))
                {
                    FileFunctions.WriteLine(_logFile, "Error calculating nearest distance from " + _tempFCOutputFile + " to " + _searchLayerName);
                    _searchErrors = true;

                    return -1;
                }
            }
            // Calculate the centroid distance and bearing if required.
            else if (includeNearFields.Equals("CENTROID", StringComparison.OrdinalIgnoreCase))
            {
                // Convert the output features to points.
                if (!await ArcGISFunctions.FeatureToPointAsync(_tempFCOutputFile, _tempFCPointsOutputFile,
                    "CENTROID", addToMap: false))
                {
                    FileFunctions.WriteLine(_logFile, "Error converting " + _tempFCOutputFile + " features to points into " + _tempFCPointsOutputFile);
                    _searchErrors = true;

                    return -1;
                }

                // Convert the search features to points.
                if (!await ArcGISFunctions.FeatureToPointAsync(_searchLayerName, _tempSearchPointsOutputFile,
                    "CENTROID", addToMap: false))
                {
                    FileFunctions.WriteLine(_logFile, "Error converting " + _searchLayerName + " features to points into " + _tempSearchPointsOutputFile);
                    _searchErrors = true;

                    return -1;
                }

                // Calculate the distance and additional proximity fields.
                if (!await ArcGISFunctions.NearAnalysisAsync(_tempFCPointsOutputFile, _tempSearchPointsOutputFile,
                    radiusText, "LOCATION", "ANGLE", "PLANAR", null, "METERS"))
                {
                    FileFunctions.WriteLine(_logFile, "Error calculating nearest distance from " + _tempFCPointsOutputFile + " to " + _tempSearchPointsOutputFile);
                    _searchErrors = true;

                    return -1;
                }

                string joinFields = "NEAR_DIST;NEAR_ANGLE";

                // Join the distance and addition proximity fields to the output feature layer.
                if (!await ArcGISFunctions.JoinFieldsAsync(_tempFCLayerName, "OBJECTID", _tempFCPointsOutputFile, "ORIG_FID",
                    joinFields, addToMap: true))
                {
                    FileFunctions.WriteLine(_logFile, "Error joining fields to " + _tempFCLayerName + " from " + _tempFCPointsOutputFile);
                    _searchErrors = true;

                    return -1;
                }
            }

            // Check the output feature layer exists in the active map.
            FeatureLayer outputFeatureLayer = await _mapFunctions.FindLayerAsync(_tempFCLayerName, null);
            if (outputFeatureLayer == null)
                return -1;

            // Check all the requested group by fields exist.
            // Only pass those that do.
            if (!string.IsNullOrEmpty(mapGroupColumns))
            {
                List<string> groupColumnList = [.. mapGroupColumns.Split(';')];
                mapGroupColumns = "";
                foreach (string groupColumn in groupColumnList)
                {
                    string columnName = groupColumn.Trim();

                    if (await _mapFunctions.FieldExistsAsync(_tempFCLayerName, columnName, null))
                        mapGroupColumns = mapGroupColumns + columnName + ";";
                }
                if (!string.IsNullOrEmpty(mapGroupColumns))
                    mapGroupColumns = mapGroupColumns.Substring(0, mapGroupColumns.Length - 1);
            }

            // Check all the requested statistics fields exist.
            // Only pass those that do.
            if (!string.IsNullOrEmpty(mapStatsColumns))
            {
                List<string> statsColumnList = [.. mapStatsColumns.Split(';')];
                mapStatsColumns = "";
                foreach (string statsColumn in statsColumnList)
                {
                    List<string> statsComponents = [.. statsColumn.Split(' ')];
                    string columnName = statsComponents[0].Trim(); // The field name.

                    if (await _mapFunctions.FieldExistsAsync(_tempFCLayerName, columnName, null))
                        mapStatsColumns = mapStatsColumns + statsColumn + ";";
                }
                if (!string.IsNullOrEmpty(mapStatsColumns))
                    mapStatsColumns = mapStatsColumns.Substring(0, mapStatsColumns.Length - 1);
            }

            // If we have group columns but no statistics columns, add a dummy column.
            if (string.IsNullOrEmpty(mapStatsColumns) && !string.IsNullOrEmpty(mapGroupColumns))
            {
                string strDummyField = mapGroupColumns.Split(';').ToList()[0];
                mapStatsColumns = strDummyField + " FIRST";
            }

            // Now do the summary statistics as required, or export the layer to table if not.
            if (!string.IsNullOrEmpty(mapStatsColumns))
            {
                FileFunctions.WriteLine(_logFile, "Calculating summary statistics ...");

                string statisticsFields = "";
                if (!string.IsNullOrEmpty(mapStatsColumns))
                    statisticsFields = mapStatsColumns;

                string caseFields = "";
                if (!string.IsNullOrEmpty(mapGroupColumns))
                    caseFields = mapGroupColumns;

                // Add the radius column to the stats columns if it's not already there.
                if (radiusText != "none")
                {
                    if (!statisticsFields.Contains("Radius FIRST", StringComparison.OrdinalIgnoreCase))
                        statisticsFields += ";Radius FIRST";
                }

                // Add the area column to the stats columns if it's not already there.
                if (!string.IsNullOrEmpty(areaColumnName))
                {
                    if (!statisticsFields.Contains(areaColumnName + " FIRST", StringComparison.OrdinalIgnoreCase))
                        statisticsFields += ";" + areaColumnName + " FIRST";
                }

                // Calculate the summary statistics.
                if (!await ArcGISFunctions.CalculateSummaryStatisticsAsync(_tempFCOutputFile, _tempTableOutputFile, statisticsFields, caseFields, addToMap: true))
                {
                    FileFunctions.WriteLine(_logFile, "Error calculating summary statistics for '" + _tempFCOutputFile + "' into " + _tempTableOutputFile);
                    _searchErrors = true;

                    return -1;
                }

                // Get the list of fields for the input table.
                IReadOnlyList<Field> inputFields;
                inputFields = await _mapFunctions.GetTableFieldsAsync(_tempTableLayerName, null);

                // Check a list of fields is returned.
                if (inputFields == null || inputFields.Count == 0)
                    return -1;

                // Now rename the radius field.
                if (radiusText != "none")
                {
                    string oldFieldName;
                    // Check the radius field by name.
                    try
                    {
                        oldFieldName = inputFields.Where(f => f.Name.Equals("FIRST_Radius", StringComparison.OrdinalIgnoreCase)).First().Name;
                    }
                    catch
                    {
                        // If not found then use the last field.
                        int intNewIndex = inputFields.Count - 1;
                        oldFieldName = inputFields[intNewIndex].Name;
                    }

                    if (!await ArcGISFunctions.RenameFieldAsync(_tempTableOutputFile, oldFieldName, "Radius"))
                    {
                        FileFunctions.WriteLine(_logFile, "Error renaming radius field in " + _tempTableLayerName);
                        _searchErrors = true;

                        return -1;
                    }
                }

                // Now rename the area field.
                if (!string.IsNullOrEmpty(areaColumnName))
                {
                    string oldFieldName;
                    // Check the area field by name.
                    try
                    {
                        oldFieldName = inputFields.Where(f => f.Name.Equals("FIRST_" + areaColumnName, StringComparison.OrdinalIgnoreCase)).First().Name;
                    }
                    catch
                    {
                        // If not found then use the last field.
                        int intNewIndex = inputFields.Count - 1;
                        oldFieldName = inputFields[intNewIndex].Name;
                    }

                    if (!await ArcGISFunctions.RenameFieldAsync(_tempTableOutputFile, oldFieldName, areaColumnName))
                    {
                        FileFunctions.WriteLine(_logFile, "Error renaming Area field in " + _tempTableLayerName);
                        _searchErrors = true;

                        return -1;
                    }
                }

                // Now export the output table.
                FileFunctions.WriteLine(_logFile, "Exporting to " + outputFormat.ToUpper() + " ...");
                intLineCount = await _mapFunctions.CopyTableToTextFileAsync(_tempTableLayerName, outputTableName, mapColumns, mapOrderColumns, ",", append, includeHeaders);
            }
            else
            {
                // Do straight copy of the feature class.
                FileFunctions.WriteLine(_logFile, "Exporting to " + outputFormat.ToUpper() + " ...");
                intLineCount = await _mapFunctions.CopyFCToTextFileAsync(_tempFCLayerName, outputTableName, mapColumns, mapOrderColumns, ",", append, includeHeaders);
            }

            return intLineCount;
        }

        /// <summary>
        /// Save the selected features from the current layer to a new layer
        /// and add it to the map if required.
        /// </summary>
        /// <param name="layerName"></param>
        /// <param name="outputFile"></param>
        /// <param name="addSelectedLayersOption"></param>
        /// <param name="layerFileName"></param>
        /// <param name="displayLabels"></param>
        /// <param name="labelClause"></param>
        /// <param name="labelColumn"></param>
        /// <returns></returns>
        private async Task<bool> KeepLayerAsync(string layerName, string outputFile, AddSelectedLayersOptions addSelectedLayersOption,
            string layerFileName, bool displayLabels, string labelClause, string labelColumn, Map targetMap)
        {
            // Add the extension to the output file name.
            outputFile += ".shp";

            bool addToMap = addSelectedLayersOption != AddSelectedLayersOptions.No;

            // Copy to a permanent file (note this is not the summarised layer).
            FileFunctions.WriteLine(_logFile, "Copying selected GIS features to " + layerName + ".shp ...");
            await ArcGISFunctions.CopyFeaturesAsync(_tempMasterLayerName, outputFile, false);

            // If the layer is to be added to the map
            if (addToMap)
            {
                // Add the layer to the required map.
                string mapName = targetMap?.Name ?? "Active Map";
                FileFunctions.WriteLine(_logFile, "Adding '" + layerName + "' to map '" + mapName + "'");

                if (!await _mapFunctions.AddLayerToMapAsync(outputFile, 0, layerName, targetMap))
                {
                    _searchErrors = true;

                    return false;
                }

                string symbologyFile = null;

                // If there is a layer file to apply.
                if (!string.IsNullOrEmpty(layerFileName))
                {
                    // Set the layer symbology to use.
                    symbologyFile = _layerPath + "\\" + layerFileName;
                }

                // Apply layer symbology and move to group layer.
                if (await SetLayerInMapAsync(layerName, symbologyFile, -1, targetMap) == null)
                {
                    _searchErrors = true;

                    return false;
                }

                // If labels are to be displayed.
                if (addSelectedLayersOption == AddSelectedLayersOptions.WithLabels && displayLabels)
                {
                    // Translate the label string.
                    if (!string.IsNullOrEmpty(labelClause))
                    {
                        try
                        {
                            List<string> labelOptions = [.. labelClause.Split('$')];
                            string labelFont = labelOptions[0].Split(':')[1];
                            double labelSize = double.Parse(labelOptions[1].Split(':')[1]);
                            int labelRed = int.Parse(labelOptions[2].Split(':')[1]);
                            int labelGreen = int.Parse(labelOptions[3].Split(':')[1]);
                            int labelBlue = int.Parse(labelOptions[4].Split(':')[1]);
                            string labelOverlap = labelOptions[5].Split(':')[1];
                            bool allowOverlap = labelOverlap.ToLower() switch
                            {
                                "allow" => true,
                                _ => false,
                            };

                            if (await _mapFunctions.LabelLayerAsync(layerName, labelColumn, labelFont, labelSize, "Normal",
                                labelRed, labelGreen, labelBlue, allowOverlap))
                                FileFunctions.WriteLine(_logFile, "Labels added to output " + layerName);
                        }
                        catch
                        {
                            FileFunctions.WriteLine(_logFile, "Error adding labels to '" + layerName + "'");
                            _searchErrors = true;

                            return false;
                        }
                    }
                    else if (!string.IsNullOrEmpty(labelColumn) && string.IsNullOrEmpty(layerFileName))
                    {
                        // Set simple labels.
                        if (await _mapFunctions.LabelLayerAsync(layerName, labelColumn))
                            FileFunctions.WriteLine(_logFile, "Labels added to output " + layerName);
                    }
                }
                else
                {
                    // Turn labels off.
                    await _mapFunctions.SwitchLabelsAsync(layerName, displayLabels);
                }
            }
            else
            {
                // User doesn't want to add the layer to the display.
                // In case it's still there from a previous run.
                await _mapFunctions.RemoveLayerAsync(layerName, targetMap);
            }

            return true;
        }

        /// <summary>
        /// Trigger the required VB macro to post-process the outputs for the
        /// current layer.
        /// </summary>
        /// <param name="macroName"></param>
        /// <param name="mapTableOutputName"></param>
        /// <param name="mapFormat"></param>
        /// <returns></returns>
        public bool StartProcess(string macroName, string mapTableOutputName, string mapFormat)
        {
            using Process scriptProc = new();

            // Set the process parameters.
            scriptProc.StartInfo.FileName = @"wscript.exe";
            scriptProc.StartInfo.WorkingDirectory = FileFunctions.GetDirectoryName(macroName);
            scriptProc.StartInfo.UseShellExecute = true;
            scriptProc.StartInfo.Arguments = string.Format("//B //Nologo \"{0}\" \"{1}\" \"{2}\" \"{3}\"", macroName, _outputPath, mapTableOutputName + "." + mapFormat.ToLower(), mapTableOutputName + ".xlsx");
            scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

            try
            {
                scriptProc.Start();
                scriptProc.WaitForExit(); // <-- Optional if you want program running until your script exits.

                int exitcode = scriptProc.ExitCode;
                if (exitcode != 0)
                {
                    FileFunctions.WriteLine(_logFile, "Error executing vbscript macro. Exit code : " + exitcode);
                    _searchErrors = true;

                    return false;
                }
            }
            catch
            {
                return false;
            }
            finally
            {
                scriptProc.Close();
            }

            return true;
        }

        #endregion Methods

        #region LayersListExpand Command

        private ICommand _mapLayersListExpandCommand;

        /// <summary>
        /// Create LayersList Expand button command.
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand MapLayersListExpandCommand
        {
            get
            {
                if (_mapLayersListExpandCommand == null)
                {
                    Action<object> expandMapLayersListAction = new(MapLayersListExpandCommandClick);
                    _mapLayersListExpandCommand = new RelayCommand(expandMapLayersListAction, param => true);
                }
                return _mapLayersListExpandCommand;
            }
        }

        /// <summary>
        /// Handles event when LayersListExpand button is pressed.
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void MapLayersListExpandCommandClick(object param)
        {
            if (_mapLayersListHeight == null)
                _mapLayersListHeight = 382;
            else
                _mapLayersListHeight = null;

            OnPropertyChanged(nameof(MapLayersListHeight));
            OnPropertyChanged(nameof(MapLayersListExpandButtonContent));
        }

        #endregion LayersListExpand Command

        #region Debugging Aides

        /// <summary>
        /// Warns the developer if this object does not have
        /// a public property with the specified name. This
        /// method does not exist in a Release build.
        /// </summary>
        [Conditional("DEBUG")]
        [DebuggerStepThrough]
        public void VerifyPropertyName(string propertyName)
        {
            // Verify that the property name matches a real,
            // public, instance property on this object.
            if (TypeDescriptor.GetProperties(this)[propertyName] == null)
            {
                string msg = "Invalid property name: " + propertyName;

                if (ThrowOnInvalidPropertyName)
                    throw new(msg);
                else
                    Debug.Fail(msg);
            }
        }

        /// <summary>
        /// Returns whether an exception is thrown, or if a Debug.Fail() is used
        /// when an invalid property name is passed to the VerifyPropertyName method.
        /// The default value is false, but subclasses used by unit tests might
        /// override this property's getter to return true.
        /// </summary>
        protected virtual bool ThrowOnInvalidPropertyName { get; private set; }

        #endregion Debugging Aides

        #region INotifyPropertyChanged Members

        /// <summary>
        /// Raised when a property on this object has a new value.
        /// </summary>
        public new event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises this object's PropertyChanged event.
        /// </summary>
        /// <param name="propertyName">The property that has a new value.</param>
        internal virtual void OnPropertyChanged(string propertyName)
        {
            VerifyPropertyName(propertyName);

            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                PropertyChangedEventArgs e = new(propertyName);
                handler(this, e);
            }
        }

        #endregion INotifyPropertyChanged Members
    }

    /// <summary>
    /// Map layers to search.
    /// </summary>
    public class MapLayer : INotifyPropertyChanged
    {
        #region Fields

        public string NodeName { get; set; }

        public string NodeGroup { get; set; }

        public string NodeLayer { get; set; }

        public string LayerName { get; set; }

        public string MapName { get; set; }

        public string GISOutputName { get; set; }

        public string TableOutputName { get; set; }

        public string Columns { get; set; }

        public string GroupColumns { get; set; }

        public string StatisticsColumns { get; set; }

        public string OrderColumns { get; set; }

        public string Criteria { get; set; }

        public bool IncludeArea { get; set; }

        public string IncludeNearFields { get; set; }

        public bool IncludeRadius { get; set; }

        public string KeyColumn { get; set; }

        public string Format { get; set; }

        public bool KeepLayer { get; set; }

        public string OutputType { get; set; }

        public bool LoadWarning { get; set; }

        public bool PreselectLayer { get; set; }

        public bool DisplayLabels { get; set; }

        public string LayerFileName { get; set; }

        public bool OverwriteLabels { get; set; }

        public string LabelColumn { get; set; }

        public string LabelClause { get; set; }

        public string MacroName { get; set; }

        public string CombinedSitesColumns { get; set; }

        public string CombinedSitesGroupColumns { get; set; }

        public string CombinedSitesStatisticsColumns { get; set; }

        public string CombinedSitesOrderByColumns { get; set; }

        private bool _isSelected;

        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                _isSelected = value;

                OnPropertyChanged(nameof(IsSelected));
            }
        }

        #endregion Fields

        #region Creator

        public MapLayer()
        {
            // constructor takes no arguments.
        }

        public MapLayer(string nodeName)
        {
            NodeName = nodeName;
        }

        #endregion Creator

        #region INotifyPropertyChanged Members

        /// <summary>
        /// Raised when a property on this object has a new value.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Raises this object's PropertyChanged event.
        /// </summary>
        /// <param name="propertyName">The property that has a new value.</param>
        internal virtual void OnPropertyChanged(string propertyName)
        {
            //VerifyPropertyName(propertyName);

            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                PropertyChangedEventArgs e = new(propertyName);
                handler(this, e);
            }
        }

        #endregion INotifyPropertyChanged Members
    }
}