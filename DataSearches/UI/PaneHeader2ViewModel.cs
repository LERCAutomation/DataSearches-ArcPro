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

using ArcGIS.Core.Data;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Internal.Framework.Controls;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
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

        private bool _updateTable;

        private string _repChar;
        private bool _requireSiteName;
        private int _defaultAddSelectedLayers;
        private int _defaultOverwriteLabels;
        private int _defaultCombinedSitesTable;

        private List<string> _bufferUnitOptionsDisplay;
        private List<string> _bufferUnitOptionsProcess;
        private List<string> _bufferUnitOptionsShort;

        private List<MapLayer> _mapLayers;

        private string _saveRootDir;
        private string _saveFolder;
        private string _gisFolder;
        private string _outputFolder;
        private string _layerFolder;

        private string _logFileName;
        private string _combinedSitesTableName;
        private string _combinedSitesColumnList;
        private string _combinedSitesTableFormat;
        private string _combinedSitesOutputFile;

        private string _bufferPrefix;
        private string _bufferLayerName;
        private string _bufferOutputFile;
        private string _bufferLayerFile;
        private string _bufferFields;
        private bool _keepBuffer;

        private string _searchLayerName;
        private string _searchOutputFile;
        private string _groupLayerName;

        private bool _keepSearchFeature;
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

        private Geodatabase _tempGDB;
        private string _tempMasterLayerName;
        private string _tempMasterOutputFile;
        private string _tempOutputLayerName;
        private string _tempFCOutputFile;
        private string _tempOutputTableName;
        private string _tempTableOutputFile;

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
            // Create a new map functions object.
            _mapFunctions = new();

            // Get the relevant config file settings.
            _repChar = _toolConfig.RepChar;

            _requireSiteName = _toolConfig.RequireSiteName;

            _defaultAddSelectedLayers = _toolConfig.DefaultAddSelectedLayers;
            _defaultOverwriteLabels = _toolConfig.DefaultOverwriteLabels;
            _defaultCombinedSitesTable = _toolConfig.DefaultCombinedSitesTable;

            _bufferUnitOptionsDisplay = _toolConfig.BufferUnitOptionsDisplay;
            _bufferUnitOptionsProcess = _toolConfig.BufferUnitOptionsProcess;
            _bufferUnitOptionsShort = _toolConfig.BufferUnitOptionsShort;

            _updateTable = _toolConfig.UpdateTable;

            _saveRootDir = _toolConfig.SaveRootDir;
            _saveFolder = _toolConfig.SaveFolder;
            _gisFolder = _toolConfig.GISFolder;
            _logFileName = _toolConfig.LogFileName;
            _layerFolder = _toolConfig.LayerFolder;
            _combinedSitesTableName = _toolConfig.CombinedSitesTableName;
            _combinedSitesColumnList = _toolConfig.CombinedSitesTableColumns;
            _combinedSitesTableFormat = _toolConfig.CombinedSitesTableFormat;

            _bufferPrefix = _toolConfig.BufferPrefix;
            _bufferLayerFile = _toolConfig.BufferLayerFile;
            _searchLayerName = _toolConfig.SearchOutputName;
            _groupLayerName = _toolConfig.GroupLayerName;

            _keepSearchFeature = _toolConfig.KeepSearchFeature;
            _searchSymbologyBase = _toolConfig.SearchSymbologyBase;

            _searchLayerBase = _toolConfig.SearchLayer;
            _searchLayerExtensions = _toolConfig.SearchLayerExtensions;
            _searchColumn = _toolConfig.SearchColumn;
            _siteColumn = _toolConfig.SiteColumn;
            _orgColumn = _toolConfig.OrgColumn;
            _radiusColumn = _toolConfig.RadiusColumn;

            _bufferFields = _toolConfig.AggregateColumns;
            _keepBuffer = _toolConfig.KeepBufferArea;

            // Get all of the map layer details.
            _mapLayers = _toolConfig.MapLayers;
        }

        #endregion Creator

        #region Controls Enabled

        /// <summary>
        /// Is the site name text box enabled.
        /// </summary>
        public bool SiteNameTextEnabled
        {
            get
            {
                return (_processingLabel == null);
                //&& (_requireSiteName) ???
            }
        }

        /// <summary>
        /// Is the list of layers enabled?
        /// </summary>
        public bool LayersListEnabled
        {
            get
            {
                return ((_processingLabel == null)
                    && (_openLayersList != null));
            }
        }

        /// <summary>
        /// Is the list of buffer units options enabled?
        /// </summary>
        public bool BufferUnitsListEnabled
        {
            get
            {
                return ((_processingLabel == null)
                    && (_bufferUnitsList != null));
            }
        }

        /// <summary>
        /// Is the list of add to map options enabled?
        /// </summary>
        public bool AddToMapListEnabled
        {
            get
            {
                return ((_processingLabel == null)
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
                return ((_processingLabel == null)
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
                return ((_processingLabel == null)
                    && (_combinedSitesList != null));
            }
        }

        /// <summary>
        /// Can the Reset button be pressed?
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool ResetButtonEnabled
        {
            get
            {
                return (_processingLabel == null);
            }
        }

        /// <summary>
        /// Can the run button be pressed?
        /// </summary>
        public bool RunButtonEnabled
        {
            get
            {
                //return true;
                return ((_processingLabel == null)
                    && (_openLayersList != null)
                    && (_openLayersList.Where(p => p.IsSelected).Any())
                    && (_searchRefText != null)
                    && (!_requireSiteName || _siteNameText != null)
                    && (!String.IsNullOrEmpty(_bufferSizeText))
                    && (_selectedBufferUnitsIndex >= 0)
                    && (_defaultAddSelectedLayers <= 0 || _selectedAddToMap != null)
                    && (_defaultOverwriteLabels <= 0 || _selectedOverwriteLabels != null)
                    && (_defaultCombinedSitesTable <= 0 || _selectedCombinedSites != null));
            }
        }

        #endregion Controls Enabled

        #region Controls Visibility

        /// <summary>
        /// Is the list of add to map options visible?
        /// </summary>
        public Visibility AddToMapListVisibility
        {
            get
            {
                if (_defaultAddSelectedLayers == -1)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        /// <summary>
        /// Is the list of overwrite options visible?
        /// </summary>
        public Visibility OverwriteLabelsListVisibility
        {
            get
            {
                if (_defaultAddSelectedLayers == -1)
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

        public Visibility HasMessage
        {
            get
            {
                if ((_processingLabel != null)
                || (string.IsNullOrEmpty(_message)))
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        public void ShowMessage(string msg, MessageType messageLevel)
        {
            MessageLevel = messageLevel;
            Message = msg;
        }

        public void ClearMessage()
        {
            Message = "";
        }

        #endregion Message

        #region Processing

        public Visibility IsProcessing
        {
            get
            {
                if (_processingLabel != null)
                    return Visibility.Visible;
                else
                    return Visibility.Hidden;
            }
        }

        private string _processingLabel;

        /// <summary>
        /// Get the processing label.
        /// </summary>
        public string ProcessingLabel
        {
            get
            {
                return _processingLabel;
            }
            set
            {
                _processingLabel = value;
                OnPropertyChanged(nameof(ProcessingLabel));
                OnPropertyChanged(nameof(IsProcessing));
            }
        }

        public void StartProcessing(string msg)
        {
            ProcessingLabel = msg;
        }

        public void StopProcessing()
        {
            ProcessingLabel = null;
        }

        #endregion Processing

        #region Reset Command

        private ICommand _resetCommand;

        /// <summary>
        /// Create Reset button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand ResetCommand
        {
            get
            {
                if (_resetCommand == null)
                {
                    Action<object> clearAction = new(ResetCommandClick);
                    _resetCommand = new RelayCommand(clearAction, param => ResetButtonEnabled);
                }
                return _resetCommand;
            }
        }

        /// <summary>
        /// Handles event when Reset button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private void ResetCommandClick(object param)
        {
            // Load the form (don't wait for the response).
            Task.Run(() => ResetForm(true));
        }

        #endregion Reset Command

        #region Run Command

        private ICommand _runCommand;

        /// <summary>
        /// Create Run button command
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public ICommand RunCommand
        {
            get
            {
                if (_runCommand == null)
                {
                    Action<object> runAction = new(RunCommandClick);
                    _runCommand = new RelayCommand(runAction, param => RunButtonEnabled);
                }

                return _runCommand;
            }
        }

        /// <summary>
        /// Handles event when Run button is clicked
        /// </summary>
        /// <param name="param"></param>
        /// <remarks></remarks>
        private async void RunCommandClick(object param)
        {
            // Site ref is mandatory.
            if (string.IsNullOrEmpty(SearchRefText))
            {
                MessageBox.Show("Please enter a search reference.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Site name is not always required.
            if (_requireSiteName && string.IsNullOrEmpty(SiteNameText))
            {
                MessageBox.Show("Please enter a location.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // At least one layer must be selected,
            if (!OpenLayersList.Where(p => p.IsSelected).Any())
            {
                MessageBox.Show("Please select at least one layer to search.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // A buffer size must be entered.
            if (string.IsNullOrEmpty(BufferSizeText))
            {
                MessageBox.Show("Please enter a buffer size.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // The buffer suze must be numeric and positive.
            bool bufferNumeric = Double.TryParse(BufferSizeText, out double bufferNumber);
            if (!bufferNumeric || bufferNumber < 0) // User either entered text or a negative number
            {
                MessageBox.Show("Please enter a positive number for the buffer size.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // A buffer unit must be selected.
            if (SelectedBufferUnitsIndex < 0)
            {
                MessageBox.Show("Please select a buffer unit.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // An add to map option must be selected (if visible).
            if (_defaultAddSelectedLayers != -1)
            {
                if (string.IsNullOrEmpty(SelectedAddToMap))
                {
                    MessageBox.Show("Please select whether layers should be added to the map.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
            }

            // A overwrite labels option must be selected (if visible).
            if (_defaultOverwriteLabels != -1 && !SelectedAddToMap.Equals("no", StringComparison.OrdinalIgnoreCase))
            {
                if (string.IsNullOrEmpty(SelectedOverwriteLabels))
                {
                    MessageBox.Show("Please select whether to overwrite labels for map layers.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
            }

            // A combined sites table option must be selected (if visible).
            if (_defaultCombinedSitesTable != -1)
            {
                if (string.IsNullOrEmpty(SelectedCombinedSites))
                {
                    MessageBox.Show("Please select whether the combined sites table should be created.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
            }

            // Update the fields and buttons in the form.
            UpdateFormControls();
            _dockPane.RefreshPanel1Buttons();

            // Run the search.
            await RunSearchAsync(); // Success not currently checked.

            // Update the fields and buttons in the form.
            UpdateFormControls();
            _dockPane.RefreshPanel1Buttons();
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

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(RunButtonEnabled));
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

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(RunButtonEnabled));
            }
        }

        /// <summary>
        /// The tool tip for the site name textbox.
        /// </summary>
        public string SiteNameTooltip
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

        private ObservableCollection<MapLayer> _openLayersList;

        /// <summary>
        /// Get the list of loaded GIS layers.
        /// </summary>
        public ObservableCollection<MapLayer> OpenLayersList
        {
            get
            {
                return _openLayersList;
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

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(RunButtonEnabled));
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

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(RunButtonEnabled));
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

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(RunButtonEnabled));
            }
        }

        private List<string> _addToMapList;

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
                OnPropertyChanged(nameof(RunButtonEnabled));
            }
        }

        private List<string> _overwriteLabelsList;

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

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(RunButtonEnabled));
            }
        }

        private List<string> _combinedSitesList;

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

                // Update the fields and buttons in the form.
                OnPropertyChanged(nameof(RunButtonEnabled));
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

        /// <summary>
        /// Get the image for the Run button.
        /// </summary>
        public static ImageSource ButtonRunImg
        {
            get
            {
                var imageSource = System.Windows.Application.Current.Resources["GenericRun16"] as ImageSource;
                return imageSource;
            }
        }

        #endregion Properties

        #region Methods

        /// <summary>
        /// Update the fields and buttons in the form.
        /// </summary>
        private void UpdateFormControls()
        {
            UpdateFormFields();
            UpdateFormButtons();
        }

        /// <summary>
        /// Update the fields in the form.
        /// </summary>
        private void UpdateFormFields()
        {
            OnPropertyChanged(nameof(SearchRefText));
            OnPropertyChanged(nameof(SiteNameText));
            OnPropertyChanged(nameof(SiteNameTextEnabled));
            OnPropertyChanged(nameof(OpenLayersList));
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
            OnPropertyChanged(nameof(SelectedCombinedSites));
            OnPropertyChanged(nameof(CombinedSitesListVisibility));
        }

        /// <summary>
        /// Update the buttons in the form.
        /// </summary>
        private void UpdateFormButtons()
        {
            OnPropertyChanged(nameof(ResetButtonEnabled));
            OnPropertyChanged(nameof(RunButtonEnabled));
        }

        /// <summary>
        /// Set all of the form fields to their default values.
        /// </summary>
        public async Task ResetForm(bool reset)
        {
            // Clear the selections first (to avoid selections being retained).
            if (_openLayersList != null)
            {
                foreach (MapLayer layer in _openLayersList)
                {
                    layer.IsSelected = false;
                }
            }

            // Search ref and site name.
            SearchRefText = null;
            SiteNameText = null;

            // Buffer size and units.
            BufferSizeText = _toolConfig.DefaultBufferSize.ToString();
            BufferUnitsList = _toolConfig.BufferUnitOptionsDisplay;
            if (_toolConfig.DefaultBufferUnit > 0)
                SelectedBufferUnitsIndex = _toolConfig.DefaultBufferUnit - 1;

            // Add layers to map.
            AddToMapList = _toolConfig.AddSelectedLayersOptions;
            if (_toolConfig.DefaultAddSelectedLayers > 0)
                SelectedAddToMap = _toolConfig.AddSelectedLayersOptions[_toolConfig.DefaultAddSelectedLayers - 1];

            // Overwrite map layers.
            OverwriteLabelsList = _toolConfig.OverwriteLabelOptions;
            if (_toolConfig.DefaultOverwriteLabels > 0)
                SelectedOverwriteLabels = _toolConfig.OverwriteLabelOptions[_toolConfig.DefaultOverwriteLabels - 1];

            // Combined sites table.
            CombinedSitesList = _toolConfig.CombinedSitesTableOptions;
            if (_toolConfig.DefaultCombinedSitesTable > 0)
                SelectedCombinedSites = _toolConfig.CombinedSitesTableOptions[_toolConfig.DefaultCombinedSitesTable - 1];

            // Log file.
            ClearLogFile = _toolConfig.DefaultClearLogFile;
            OpenLogFile = _toolConfig.DefaultOpenLogFile;

            // Reload the form layers (don't wait for the response).
            await LoadLayersAsync(reset, true);
        }

        /// <summary>
        /// Load the list of open GIS layers.
        /// </summary>
        /// <param name="selectedTable"></param>
        /// <returns></returns>
        public async Task LoadLayersAsync(bool reset, bool message)
        {
            _dockPane.LayersListLoading = true;
            if (reset)
                StartProcessing("Resetting form ...");
            else
                StartProcessing("Loading form ...");

            // Clear any messages.
            ClearMessage();

            // Update the fields and buttons in the form.
            UpdateFormControls();

            //await Task.Delay(100);  // Temporary sleep to check processing label and animation are working. ???

            List<string> closedLayers = []; // The closed layers by name.

            await Task.Run(() =>
            {
                if (_mapFunctions == null || _mapFunctions.MapName == null || MapView.Active.Map.Name != _mapFunctions.MapName)
                {
                    // Create a new map functions object.
                    _mapFunctions = new();
                }

                // Check if there is an active map.
                bool mapOpen = (_mapFunctions.MapName != null);

                // Reset the list of open layers.
                ObservableCollection<MapLayer> openLayersList = [];

                if (mapOpen)
                {
                    List<MapLayer> allLayers = _mapLayers;

                    // Loop through all of the layers to check if they are open
                    // in the active map.
                    foreach (MapLayer layer in allLayers)
                    {
                        if (_mapFunctions.FindLayer(layer.LayerName) != null)
                        {
                            // Preselect layer if required.
                            layer.IsSelected = layer.PreselectLayer;

                            // Add the open layers to the list.
                            openLayersList.Add(layer);
                        }
                        else
                        {
                            // Only add if the user wants to be warned of this one.
                            if (layer.LoadWarning)
                                closedLayers.Add(layer.LayerName);
                        }
                    }
                }

                // Reset the list of open layers.
                _openLayersList = openLayersList;
            });

            StopProcessing();

            // Show a message if there are no open layers.
            if (!_openLayersList.Any())
                ShowMessage("No search layers in active map", MessageType.Warning);

            _dockPane.LayersListLoading = false;

            // Update the fields and buttons in the form.
            UpdateFormControls();

            // Warn the user of closed layers.
            int closedLayerCount = closedLayers.Count;
            if (closedLayerCount > 0)
            {
                string closedLayerWarning = "";
                if (closedLayerCount == 1)
                {
                    closedLayerWarning = "Warning. The layer '" + closedLayers[0] + "' is not loaded.";
                }
                else if (closedLayerCount > 10)
                {
                    closedLayerWarning = String.Format("Warning: {0} layers are not loaded, including:{1}{1}", closedLayerCount.ToString(), Environment.NewLine);
                    closedLayerWarning += String.Join(Environment.NewLine, closedLayers.Take(10));
                }
                else
                {
                    closedLayerWarning = String.Format("Warning: {0} layers are not loaded:{1}{1}", closedLayerCount.ToString(), Environment.NewLine);
                    closedLayerWarning += String.Join(Environment.NewLine, closedLayers);
                }

                if (message)
                    MessageBox.Show(closedLayerWarning, "Data Searches", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>
        /// Clear the list of open GIS layers.
        /// </summary>
        /// <param name="selectedTable"></param>
        /// <returns></returns>
        public void ClearLayers()
        {
            // Clear the list of open layers.
            _openLayersList = [];

            // Update the fields and buttons in the form.
            UpdateFormControls();
        }

        /// <summary>
        /// Validate and run the search.
        /// </summary>
        private async Task RunSearchAsync()
        {
            if (_mapFunctions == null || _mapFunctions.MapName == null)
            {
                // Create a new map functions object.
                _mapFunctions = new();
            }

            // Save the parameters.
            string searchRef = SearchRefText;
            string siteName = SiteNameText;
            //string strOrganisation = txtOrganisation.Text; // Associated organisation
            string bufferSize = BufferSizeText;
            int bufferUnitIndex = SelectedBufferUnitsIndex;

            // Selected layers.
            _selectedLayers = OpenLayersList.Where(p => p.IsSelected).ToList();

            // What is the selected buffer unit?
            string bufferUnitText = _bufferUnitOptionsDisplay[bufferUnitIndex]; // Unit to be used in reporting.
            string bufferUnitProcess = _bufferUnitOptionsProcess[bufferUnitIndex]; // Unit to be used in process (because of American spelling).
            string bufferUnitShort = _bufferUnitOptionsShort[bufferUnitIndex]; // Unit to be used in file naming (abbreviation).

            // What is the area measurement unit?
            string areaMeasureUnit = _toolConfig.AreaMeasurementUnit;

            // Will selected layers be added to the map with labels?
            AddSelectedLayersOptions addSelectedLayersOption = AddSelectedLayersOptions.No;
            if (_defaultAddSelectedLayers > 0)
            {
                if (SelectedAddToMap.Equals("no", StringComparison.CurrentCultureIgnoreCase))
                    addSelectedLayersOption = AddSelectedLayersOptions.No;
                else if (SelectedAddToMap.Contains("with labels", StringComparison.CurrentCultureIgnoreCase))
                    addSelectedLayersOption = AddSelectedLayersOptions.WithLabels;
                else if (SelectedAddToMap.Contains("without labels", StringComparison.CurrentCultureIgnoreCase))
                    addSelectedLayersOption = AddSelectedLayersOptions.WithoutLabels;
            }

            // Will the labels on map layers be overwritten?
            OverwriteLabelOptions overwriteLabelsOption = OverwriteLabelOptions.No;
            if (_defaultOverwriteLabels > 0)
            {
                if (SelectedOverwriteLabels.Equals("no", StringComparison.CurrentCultureIgnoreCase))
                    overwriteLabelsOption = OverwriteLabelOptions.ResetByLayer;
                else if (SelectedOverwriteLabels.Contains("reset each layer", StringComparison.CurrentCultureIgnoreCase))
                    overwriteLabelsOption = OverwriteLabelOptions.ResetByLayer;
                else if (SelectedOverwriteLabels.Contains("reset each group", StringComparison.CurrentCultureIgnoreCase))
                    overwriteLabelsOption = OverwriteLabelOptions.ResetByGroup;
                else if (SelectedOverwriteLabels.Contains("do not reset", StringComparison.CurrentCultureIgnoreCase))
                    overwriteLabelsOption = OverwriteLabelOptions.DoNotReset;
            }

            // Will the combined sites table be created and overwritten?
            CombinedSitesTableOptions combinedSitesTableOption = CombinedSitesTableOptions.None;
            if (_defaultCombinedSitesTable > 0)
            {
                if (SelectedCombinedSites.Equals("none", StringComparison.CurrentCultureIgnoreCase))
                    combinedSitesTableOption = CombinedSitesTableOptions.Append;
                else if (SelectedCombinedSites.Contains("append", StringComparison.CurrentCultureIgnoreCase))
                    combinedSitesTableOption = CombinedSitesTableOptions.Append;
                else if (SelectedCombinedSites.Contains("overwrite", StringComparison.CurrentCultureIgnoreCase))
                    combinedSitesTableOption = CombinedSitesTableOptions.Append;
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
            _saveFolder = StringFunctions.ReplaceSearchStrings(_saveFolder, reference, siteName, shortRef, subref, radius);
            _gisFolder = StringFunctions.ReplaceSearchStrings(_gisFolder, reference, siteName, shortRef, subref, radius);
            _logFileName = StringFunctions.ReplaceSearchStrings(_logFileName, reference, siteName, shortRef, subref, radius);
            _combinedSitesTableName = StringFunctions.ReplaceSearchStrings(_combinedSitesTableName, reference, siteName, shortRef, subref, radius);
            _bufferPrefix = StringFunctions.ReplaceSearchStrings(_bufferPrefix, reference, siteName, shortRef, subref, radius);
            _searchLayerName = StringFunctions.ReplaceSearchStrings(_searchLayerName, reference, siteName, shortRef, subref, radius);
            _groupLayerName = StringFunctions.ReplaceSearchStrings(_groupLayerName, reference, siteName, shortRef, subref, radius);

            // Remove any illegal characters from the names.
            _saveFolder = StringFunctions.StripIllegals(_saveFolder, _repChar);
            _gisFolder = StringFunctions.StripIllegals(_gisFolder, _repChar);
            _logFileName = StringFunctions.StripIllegals(_logFileName, _repChar, true);
            _combinedSitesTableName = StringFunctions.StripIllegals(_combinedSitesTableName, _repChar);
            _bufferPrefix = StringFunctions.StripIllegals(_bufferPrefix, _repChar);
            _searchLayerName = StringFunctions.StripIllegals(_searchLayerName, _repChar);
            _groupLayerName = StringFunctions.StripIllegals(_groupLayerName, _repChar);

            // Trim any trailing spaces (directory functions don't deal with them well).
            _saveFolder = _saveFolder.Trim();

            // Create output folders if required.
            _outputFolder = CreateOutputFolders(_saveRootDir, _saveFolder, _gisFolder);
            if (_outputFolder == null)
            {
                //TODO ???
                return;
            }

            // Create log file (if necessary).
            _logFile = _outputFolder + @"\" + _logFileName;
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
                    return;
                }
            }

            // Replace any illegal characters in the user name string.
            _userID = StringFunctions.StripIllegals(Environment.UserName, "_", false);

            // User ID should be something at least.
            if (string.IsNullOrEmpty(_userID))
            {
                _userID = "Temp";
                FileFunctions.WriteLine(_logFile, "User ID not found. User ID used will be 'Temp'");
            }

            // Indicate the search has started.
            _dockPane.SearchRunning = true;
            StartProcessing("Performing search ...");

            // Temporary delay to check processing label and animation are working. ???
            //await Task.Delay(1000); ???

            // Write the first line to the log file.
            FileFunctions.WriteLine(_logFile, "-----------------------------------------------------------------------");
            FileFunctions.WriteLine(_logFile, "Processing search '" + reference + "'");
            FileFunctions.WriteLine(_logFile, "-----------------------------------------------------------------------");

            FileFunctions.WriteLine(_logFile, "Parameters are as follows:");
            FileFunctions.WriteLine(_logFile, "Buffer distance: " + radius);
            FileFunctions.WriteLine(_logFile, "Output location: " + _saveFolder);
            //TODO ???
            //string strLayerString = "";
            //foreach (string aLayer in SelectedLayers)
            //    strLayerString = strLayerString + aLayer + ", ";
            FileFunctions.WriteLine(_logFile, "Layers to process: " + SelectedLayers.Count.ToString());
            FileFunctions.WriteLine(_logFile, "Area measurement unit: " + areaMeasureUnit);

            // Create the search query, e.g. "TERM = 'City'".
            string searchClause = _searchColumn + " = '" + reference + "'";

            // Count the features matching the search reference.
            long featureCount = await CountSearchFeaturesAsync(searchClause);
            if (featureCount == 0)
            {
                //TODO ???
                StopSearch(searchRef, false, false);
                return;
            }

            // Prepare the temporary geodatabase
            if (!await PrepareTemporaryGDBAsync())
            {
                //TODO ???
                StopSearch(searchRef, false, true);
                return;
            }

            // Select the feature matching the search reference in the map.
            if (!await _mapFunctions.SelectLayerByAttributesAsync(_inputLayerName, searchClause, SelectionCombinationMethod.New))
            {
                //TODO ???
                StopSearch(searchRef, false, true);
                return;
            }

            //TODO ???
            // Update the table if required.
            //UpdateSearchFeaturesAsync();

            // The output file for the search features is a shapefile in the root save directory.
            _searchOutputFile = _outputFolder + "\\" + _searchLayerName + ".shp";

            // Save the selected feature(s).
            if (!await SaveSearchFeaturesAsync())
            {
                MessageBox.Show("Error saving search feature(s).");
                FileFunctions.WriteLine(_logFile, "Error saving search feature(s)");

                StopSearch(searchRef, false, false);
                return;
            }

            // Set the buffer layer name by appending the radius.
            _bufferLayerName = _bufferPrefix + "_" + radius;
            if (_bufferLayerName.Contains('.'))
                _bufferLayerName = _bufferLayerName.Replace('.', '_');

            // The output file for the buffer is a shapefile in the root save directory.
            _bufferOutputFile = _outputFolder + "\\" + _bufferLayerName + ".shp";

            // Buffer search feature(s).
            if (!await BufferSearchFeaturesAsync(bufferSize, bufferUnitProcess))
            {
                //TODO ???
                StopSearch(searchRef, false, true);
                return;
            }

            // Start the combined sites table before we do any analysis.
            _combinedSitesOutputFile = _outputFolder + @"\" + _combinedSitesTableName + "." + _combinedSitesTableFormat;
            if (!CreateCombinedSitesTable(_combinedSitesOutputFile, combinedSitesTableOption))
            {
                //TODO ???
                StopSearch(searchRef, false, true);
                return;
            }

            // Loop through the map layers, processing each one.
            await ProcessMapLayersAsync(reference, siteName, shortRef, subref, radius, areaMeasureUnit, addSelectedLayersOption, overwriteLabelsOption, combinedSitesTableOption);

            FileFunctions.WriteLine(_logFile, "");

            // Clean up after the search
            await CleanUpSearchAsync(addSelectedLayersOption);

            // Indicate that the search process has completed successfully.
            StopSearch(searchRef, true, true);

            return;
        }

        /// <summary>
        /// Indicate that the search process has stopped (either
        /// successfully or otherwise).
        /// </summary>
        /// <param name="searchRef"></param>
        /// <param name="success"></param>
        /// <param name="message"></param>
        private void StopSearch(string searchRef, bool success, bool message)
        {
            // Indicate search has finished.
            _dockPane.SearchRunning = false;
            StopProcessing();

            // Delete the group layer if it is empty.   ???

            if (success)
            {
                FileFunctions.WriteLine(_logFile, "---------------------------------------------------------------------------");
                FileFunctions.WriteLine(_logFile, "Process complete");
                FileFunctions.WriteLine(_logFile, "---------------------------------------------------------------------------");

                if (message)
                {
                    Notification notification = new()
                    {
                        Title = FrameworkApplication.Title,
                        Severity = Notification.SeverityLevel.High,
                        Message = String.Format("Search '{0}' complete!", searchRef),
                        ImageSource = System.Windows.Application.Current.Resources["ToastLicensing32"] as ImageSource   // ???
                    };
                    FrameworkApplication.AddNotification(notification);

                    //MessageBox.Show(String.Format("Search '{0}' complete!", searchRef), "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            else
            {
                FileFunctions.WriteLine(_logFile, "Process aborted");

                if (message)
                {
                    Notification notification = new()
                    {
                        Title = FrameworkApplication.Title,
                        Severity = Notification.SeverityLevel.High,
                        Message = String.Format("Search '{0}' aborted!", searchRef),
                        ImageSource = System.Windows.Application.Current.Resources["ToastLicensing32"] as ImageSource   // ???
                    };
                    FrameworkApplication.AddNotification(notification);

                    //MessageBox.Show(String.Format("Search '{0}' aborted!", searchRef), "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }

            // Open the log file (if required).
            if (OpenLogFile && message)
                Process.Start("notepad.exe", _logFile);
        }

        private async Task CleanUpSearchAsync(AddSelectedLayersOptions addSelectedLayersOption)
        {
            // Save the selected feature if required.
            if (_keepSearchFeature)
            {
                // Add the search layer to the map.
                if (addSelectedLayersOption != AddSelectedLayersOptions.No)
                {
                    // Set the search layer symbology to use.
                    string searchlayerFile = _searchSymbologyBase + _searchLayerExtension + ".lyrx";
                    string symbologyFile = _layerFolder + "\\" + searchlayerFile;

                    if (!await SetLayerInMapAsync(_searchLayerName, symbologyFile))
                    {
                        MessageBox.Show("Error setting search feature layer in the map.");
                        FileFunctions.WriteLine(_logFile, "Error setting search feature layer in the map");
                    }

                    FileFunctions.WriteLine(_logFile, "Search feature layer added to display");
                }
                else
                {
                    // Remove the search feature layer from the map.
                    await _mapFunctions.RemoveLayerAsync(_searchLayerName);
                }
            }
            else
            {
                try
                {
                    // Remove the search feature layer from the map.
                    await _mapFunctions.RemoveLayerAsync(_searchLayerName);

                    // Delete the search feature class.
                    string searchOutputFile = _outputFolder + "\\" + _searchLayerName + ".shp";
                    await ArcGISFunctions.DeleteFeatureClassAsync(searchOutputFile);

                    FileFunctions.WriteLine(_logFile, "Search feature layer deleted");
                }
                catch
                {
                    //MessageBox.Show("Error deleting the search feature layer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Do we want to keep the buffer layer? If not, remove it.
            if (_keepBuffer)
            {
                // Add the output buffer layer to the map.
                if (addSelectedLayersOption != AddSelectedLayersOptions.No)
                {
                    // Set the buffer layer symbology to use.
                    string symbologyFile = _layerFolder + "\\" + _bufferLayerFile;

                    if (!await SetLayerInMapAsync(_bufferLayerName, symbologyFile))
                    {
                        MessageBox.Show("Error setting buffer layer in the map.");
                        FileFunctions.WriteLine(_logFile, "Error setting buffer layer in the map");
                    }

                    FileFunctions.WriteLine(_logFile, "Buffer layer added to display");
                }
                else
                {
                    // Remove the buffer layer from the map.
                    await _mapFunctions.RemoveLayerAsync(_bufferLayerName);
                }
            }
            else
            {
                try
                {
                    // Remove the buffer layer from the map.
                    await _mapFunctions.RemoveLayerAsync(_bufferLayerName);

                    // Delete the buffer feature class.
                    await ArcGISFunctions.DeleteFeatureClassAsync(_bufferOutputFile);

                    FileFunctions.WriteLine(_logFile, "Buffer layer deleted");
                }
                catch
                {
                    //MessageBox.Show("Error deleting the buffer layer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // Remove the temporary master layer from the map.
            await _mapFunctions.RemoveLayerAsync(_tempMasterLayerName);

            // Delete the temporary master feature class.
            await ArcGISFunctions.DeleteFeatureClassAsync(_tempMasterOutputFile);

            // Clear the search features selection.
            await _mapFunctions.ClearLayerSelectionAsync(_searchLayerName);

            // Remove the group layer from the map if it is empty.
            await _mapFunctions.RemoveGroupLayerAsync(_groupLayerName);
        }

        /// <summary>
        /// Create the output folders if required.
        /// </summary>
        /// <param name="saveRootDir"></param>
        /// <param name="saveFolder"></param>
        /// <param name="gisFolder"></param>
        /// <returns></returns>
        private static string CreateOutputFolders(string saveRootDir, string saveFolder, string gisFolder)
        {
            // Create root folder if required.
            if (!FileFunctions.DirExists(saveRootDir))
            {
                try
                {
                    Directory.CreateDirectory(saveRootDir);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Cannot create directory '" + saveRootDir + "'. System error: " + ex.Message);
                    return null;
                }
            }

            // Create save sub-folder if required.
            if (saveFolder != "")
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
                    MessageBox.Show("Cannot create directory '" + saveFolder + "'. System error: " + ex.Message);
                    return null;
                }
            }

            // Create gis sub-folder if required.
            if (gisFolder != "")
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
                    MessageBox.Show("Cannot create directory '" + gisFolder + "'. System error: " + ex.Message);
                    return null;
                }
            }

            return gisFolder;
        }

        private bool CreateCombinedSitesTable(string combinedSitesTable, CombinedSitesTableOptions combinedSitesTableOption)
        {
            // Start the table if overwrite has been selected, or if the table doesn't exist (and append has been selected).
            if ((combinedSitesTableOption == CombinedSitesTableOptions.Overwrite) ||
                (combinedSitesTableOption == CombinedSitesTableOptions.Append && !FileFunctions.FileExists(combinedSitesTable)))
            {
                if (!FileFunctions.WriteEmptyTextFile(combinedSitesTable, _toolConfig.CombinedSitesTableColumns))
                {
                    MessageBox.Show("Error writing to combined sites table.");
                    FileFunctions.WriteLine(_logFile, "Error writing to combined sites table");

                    return false;
                }

                FileFunctions.WriteLine(_logFile, "Combined sites table started");
            }

            return true;
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

                // Find the feature layer by name if it exists. Only search existing layers.
                FeatureLayer featureLayer = _mapFunctions.FindLayer(searchLayer);

                if (featureLayer != null)
                {
                    // Count the required features in the layer.
                    long featureCount = await ArcGISFunctions.CountFeaturesAsync(featureLayer, searchClause);

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
                MessageBox.Show("No features found in any of the search layers; Process aborted.");
                FileFunctions.WriteLine(_logFile, "No features found in any of the search layers");
                return 0;
            }

            // If features found in more than one layer.
            if (featureLayerCount > 1)
            {
                MessageBox.Show(totalFeatureCount.ToString() + " features found in different search layers; Process aborted.");
                FileFunctions.WriteLine(_logFile, totalFeatureCount.ToString() + " features found in different search layers");
                return 0;
            }

            // If multiple features found.
            if (totalFeatureCount > 1)
            {
                // Ask the user if they want to continue
                MessageBoxResult response = MessageBox.Show(totalFeatureCount.ToString() + " features found in " + _inputLayerName + " matching those criteria. Do you wish to continue?", "Data Searches", MessageBoxButton.YesNo);
                if (response == MessageBoxResult.No)
                {
                    FileFunctions.WriteLine(_logFile, totalFeatureCount.ToString() + " features found in the search layers");
                    FileFunctions.WriteLine(_logFile, "Process aborted");
                    return 0;
                }
            }

            return totalFeatureCount;
        }

        private async Task<bool> PrepareTemporaryGDBAsync()
        {
            // Set a temporary folder path.
            string tempFolder = System.IO.Path.GetTempPath();

            // Create the temporary file geodatabase if it doesn't exist.
            _tempGDBName = tempFolder + @"Temp.gdb";
            _tempGDB = null;
            if (!FileFunctions.DirExists(_tempGDBName))
            {
                _tempGDB = ArcGISFunctions.CreateFileGeodatabase(_tempGDBName);
                if (_tempGDB == null)
                {
                    MessageBox.Show("Error creating temporary geodatabase" + _tempGDBName);
                    FileFunctions.WriteLine(_logFile, "Error creating temporary geodatabase " + _tempGDBName);
                    return false;
                }

                FileFunctions.WriteLine(_logFile, "Temporary geodatabase created");
            }

            // Delete any temporary feature classes and tables that still exist
            _tempMasterLayerName = "TempMaster_" + _userID;
            _tempMasterOutputFile = _tempGDBName + @"\" + _tempMasterLayerName;
            await _mapFunctions.RemoveLayerAsync(_tempMasterLayerName);
            if (await ArcGISFunctions.FeatureClassExistsAsync(_tempMasterOutputFile))
            {
                ArcGISFunctions.DeleteFeatureClass(_tempGDBName, _tempMasterLayerName);
                FileFunctions.WriteLine(_logFile, "Temporary master feature class deleted");
            }

            _tempOutputLayerName = "TempOutput_" + _userID;
            _tempFCOutputFile = _tempGDBName + @"\" + _tempOutputLayerName;
            await _mapFunctions.RemoveLayerAsync(_tempOutputLayerName);
            if (await ArcGISFunctions.FeatureClassExistsAsync(_tempFCOutputFile))
            {
                ArcGISFunctions.DeleteFeatureClass(_tempGDBName, _tempOutputLayerName);
                FileFunctions.WriteLine(_logFile, "Temporary output feature class deleted");
            }

            _tempOutputTableName = "TempOutput_" + _userID + "DBF";
            _tempTableOutputFile = _tempGDBName + @"\" + _tempOutputTableName;
            await _mapFunctions.RemoveLayerAsync(_tempOutputTableName);
            if (await ArcGISFunctions.TableExistsAsync(_tempTableOutputFile))
            {
                ArcGISFunctions.DeleteTable(_tempGDBName, _tempOutputTableName);
                FileFunctions.WriteLine(_logFile, "Temporary output table deleted");
            }

            return true;
        }

        private async Task<bool> UpdateSearchFeaturesAsync()
        {
            //TODO ???
            return true;
        }

        private async Task<bool> SaveSearchFeaturesAsync()
        {
            // Get the full layer path (in case it's nested in one or more groups).
            string inputLayerPath = _mapFunctions.GetLayerPath(_inputLayerName);

            FileFunctions.WriteLine(_logFile, "Saving search feature(s)");

            // Copy the selected feature(s) to an output file.
            if (!await ArcGISFunctions.CopyFeaturesAsync(inputLayerPath, _searchOutputFile, true))
            {
                MessageBox.Show("Error saving search feature(s)");
                FileFunctions.WriteLine(_logFile, "Error saving search feature(s)");
                return false;
            }

            return true;
        }

        private async Task<bool> BufferSearchFeaturesAsync(string bufferSize, string bufferUnitProcess)
        {
            // Get the full layer path (in case it's nested in one or more groups).
            string searchLayerPath = _mapFunctions.GetLayerPath(_searchLayerName);

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
                if (await _mapFunctions.FieldExistsAsync(searchLayerPath, fieldName))
                {
                    dissolveFields = dissolveFields + fieldName + ";";
                }
            }

            FileFunctions.WriteLine(_logFile, "Buffering feature(s) with a distance of " + bufferDistance);

            string dissolveOption = "ALL";
            if (dissolveFields != "")
            {
                dissolveFields = dissolveFields.Substring(0, dissolveFields.Length - 1);
                dissolveOption = "LIST";
            }

            if (!await ArcGISFunctions.BufferFeaturesAsync(searchLayerPath, _bufferOutputFile, bufferDistance, "FULL", "ROUND", dissolveOption, dissolveFields, addToMap: true))
            {
                MessageBox.Show("Error during feature buffering. Process aborted");
                FileFunctions.WriteLine(_logFile, "Error during feature buffering. Process aborted");
                return false;
            }

            return true;
        }

        private async Task<bool> SetLayerInMapAsync(string layerName, string symbologyFile)
        {
            // Apply layer symbology.
            if (!String.IsNullOrEmpty(symbologyFile))
            {
                if (!await _mapFunctions.ApplySymbologyFromLayerFileAsync(layerName, symbologyFile))
                {
                    MessageBox.Show("Error applying symbology to '" + layerName + "'");
                    FileFunctions.WriteLine(_logFile, "Error applying symbology to '" + layerName + "'");
                }
            }

            // Move layer to the group layer.
            if (String.IsNullOrEmpty(_groupLayerName))
            {
                if (!await _mapFunctions.MoveToGroupLayerAsync(_mapFunctions.FindLayer(layerName), _groupLayerName, -1))
                {
                    MessageBox.Show("Error moving layer to '" + layerName + "'");
                    FileFunctions.WriteLine(_logFile, "Error moving layer to '" + layerName + "'");
                }
            }

            return true;
        }

        private async Task<bool> ProcessMapLayersAsync(string reference, string siteName, string shortRef, string subref, string radius, string areaMeasureUnit,
            AddSelectedLayersOptions addSelectedLayersOption,
            OverwriteLabelOptions overwriteLabelOption,
            CombinedSitesTableOptions combinedSitesTableOption)
        {
            // Get any groups and initialise required layers.
            if (overwriteLabelOption == OverwriteLabelOptions.ResetByLayer ||
                overwriteLabelOption == OverwriteLabelOptions.ResetByGroup)
            {
                _mapGroupNames = _selectedLayers.Select(l => l.LayerGroup).Distinct().ToList();
                foreach (string groupName in _mapGroupNames)
                {
                    _mapGroupLabels.Add(1); // each group has its own label counter.
                }
            }

            // Keep track of the label numbers.
            int startLabel = 1;
            _maxLabel = 1;

            // Get the full layer path (in case it's nested in one or more groups).
            string bufferLayerPath = _mapFunctions.GetLayerPath(_bufferLayerName);

            foreach (MapLayer selectedLayer in SelectedLayers)
            {
                // Get all the settings relevant to this layer.
                string mapGroupName = selectedLayer.LayerGroup;
                string mapLayerName = selectedLayer.LayerName;
                string mapOutputName = selectedLayer.GISOutputName;
                string mapTableOutputName = selectedLayer.TableOutputName;
                string mapColumns = selectedLayer.Columns;
                string mapGroupColumns = selectedLayer.GroupColumns;
                string mapStatsColumns = selectedLayer.StatisticsColumns;
                string mapOrderColumns = selectedLayer.OrderColumns;
                string mapCriteria = selectedLayer.Criteria;

                bool mapIncludeArea = selectedLayer.IncludeArea;
                bool mapIncludeDistance = selectedLayer.IncludeDistance;
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
                mapOutputName = StringFunctions.ReplaceSearchStrings(mapOutputName, reference, siteName, shortRef, subref, radius);
                mapTableOutputName = StringFunctions.ReplaceSearchStrings(mapTableOutputName, reference, siteName, shortRef, subref, radius);

                // Remove any illegal characters from the names.
                mapOutputName = StringFunctions.StripIllegals(mapOutputName, _repChar);
                mapTableOutputName = StringFunctions.StripIllegals(mapTableOutputName, _repChar);

                mapStatsColumns = StringFunctions.AlignStatsColumns(mapColumns, mapStatsColumns, mapGroupColumns);
                mapCombinedSitesStatsColumns = StringFunctions.AlignStatsColumns(mapCombinedSitesColumns, mapCombinedSitesStatsColumns, mapCombinedSitesGroupColumns);

                FileFunctions.WriteLine(_logFile, "");
                FileFunctions.WriteLine(_logFile, "Starting analysis for " + selectedLayer.LayerName);

                // Get the full layer path (in case it's nested in one or more groups).
                string mapLayerPath = _mapFunctions.GetLayerPath(mapLayerName);

                // Select by location.
                FileFunctions.WriteLine(_logFile, "Selecting features by location");
                if (!await ArcGISFunctions.SelectLayerByLocationAsync(mapLayerPath, bufferLayerPath, "INTERSECT", "", "NEW_SELECTION"))
                {
                    MessageBox.Show("Error selecting layer " + mapLayerName + " by location.");
                    FileFunctions.WriteLine(_logFile, "Error selecting layer " + mapLayerName + " by location");

                    return false;
                }

                // Find the map layer by name.
                FeatureLayer mapLayer = _mapFunctions.FindLayer(mapLayerName);

                // Refine the selection by attributes (if required).
                if (mapLayer.SelectionCount > 0 && mapCriteria != "")
                {
                    FileFunctions.WriteLine(_logFile, "Refining selection with criteria " + mapCriteria + " ...");

                    if (!await _mapFunctions.SelectLayerByAttributesAsync(mapLayerName, mapCriteria, SelectionCombinationMethod.And))
                    {
                        MessageBox.Show("Error selecting layer " + mapLayerName + " with criteria " + mapCriteria + ". Please check syntax and column names (case sensitive).");
                        FileFunctions.WriteLine(_logFile, "Error refining selection on layer " + mapLayerName + " with criteria " + mapCriteria + ". Please check syntax and column names (case sensitive)");

                        return false;
                    }
                }

                // Count the selected features.
                int featureCount = mapLayer.SelectionCount;

                // Write out the results - to a feature class initially. Include distance if required.
                if (featureCount > 0)
                {
                    FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", featureCount) + " feature(s) found");

                    // Create the map output depending on the output type required.
                    if (!await CreateMapOutputAsync(mapLayerName, mapLayerPath, bufferLayerPath, mapOutputType))
                    {
                        MessageBox.Show("Cannot output selection from " + mapLayerName + " to " + _tempMasterOutputFile + ".");
                        FileFunctions.WriteLine(_logFile, "Cannot output selection from " + mapLayerName + " to " + _tempMasterOutputFile);

                        return false;
                    }

                    // Add map labels to the output if required.
                    if (addSelectedLayersOption == AddSelectedLayersOptions.WithLabels && !String.IsNullOrEmpty(mapLabelColumn))
                    {
                        if (!await AddMapLabelsAsync(overwriteLabelOption, mapOverwriteLabels, mapLabelColumn, mapKeyColumn, mapGroupName))
                        {
                            MessageBox.Show("Error adding map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile + ".");
                            FileFunctions.WriteLine(_logFile, "Error adding map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile);

                            return false;
                        }
                    }

                    // Create relevant output names.
                    string mapOutputFile = _outputFolder + @"\" + mapOutputName; // Output shapefile / feature class name. Note no extension to allow write to GDB.
                    string mapTableOutputFile = _outputFolder + @"\" + mapTableOutputName + "." + mapFormat.ToLower(); // Output table name.

                    // Include headers for CSV files.
                    bool includeHeaders = false;
                    if (mapFormat.Equals("csv", StringComparison.CurrentCultureIgnoreCase))
                        includeHeaders = true;

                    // Only include radius if requested.
                    string radiusText = "none";
                    if (mapIncludeRadius)
                        radiusText = radius;

                    string areaUnit = "";
                    if (mapIncludeArea)
                        areaUnit = areaMeasureUnit;

                    // Export results to table if required.
                    if (!String.IsNullOrEmpty(mapFormat) && !String.IsNullOrEmpty(mapColumns))
                    {
                        FileFunctions.WriteLine(_logFile, "Extracting summary output ...");

                        int intLineCount = -999;
                        intLineCount = await ExportSelectionAsync(mapTableOutputFile, mapColumns, mapGroupColumns, mapStatsColumns, mapOrderColumns,
                            includeHeaders, areaUnit, mapIncludeDistance, radiusText);
                        if (intLineCount < 0)
                        {
                            MessageBox.Show("Error extracting summary from " + _tempMasterOutputFile + ".");
                            FileFunctions.WriteLine(_logFile, "Error extracting summary from " + _tempMasterOutputFile);

                            return false;
                        }

                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " record(s) exported");
                    }

                    // Copy to permanent layer as appropriate
                    if (mapKeepLayer)
                    {
                        if (!await KeepLayerAsync(mapOutputName, mapOutputFile, addSelectedLayersOption, mapLayerFileName, mapDisplayLabels, mapLabelClause, mapLabelColumn))
                        {
                            MessageBox.Show("Error saving layer " + _tempMasterOutputFile + ".");
                            FileFunctions.WriteLine(_logFile, "Error saving layer " + _tempMasterOutputFile);

                            return false;
                        }
                    }

                    // Add to combined sites table if required.
                    if (mapCombinedSitesColumns != "" && combinedSitesTableOption != CombinedSitesTableOptions.None)
                    {
                        FileFunctions.WriteLine(_logFile, "Extracting summary output for combined sites table ...");

                        int intLineCount = -999;
                        intLineCount = await ExportSelectionAsync(_combinedSitesOutputFile, mapCombinedSitesColumns, mapCombinedSitesGroupColumns,
                            mapCombinedSitesStatsColumns, mapCombinedSitesOrderColumns,
                            false, areaUnit, mapIncludeDistance, radiusText);

                        if (intLineCount < 0)
                        {
                            MessageBox.Show("Error extracting summary from " + _tempMasterOutputFile + ".");
                            FileFunctions.WriteLine(_logFile, "Error extracting summary from " + _tempMasterOutputFile);

                            return false;
                        }

                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " row(s) added to combined sites table");
                    }

                    // Cleanup the temporary master layer.
                    await _mapFunctions.RemoveLayerAsync(_tempMasterLayerName);
                    await ArcGISFunctions.DeleteFeatureClassAsync(_tempMasterOutputFile);

                    // Clear the selection in the input layer.
                    await _mapFunctions.ClearLayerSelectionAsync(mapLayerName);

                    FileFunctions.WriteLine(_logFile, "Analysis complete");
                }
                else
                {
                    FileFunctions.WriteLine(_logFile, "No features found");
                }
            }

            return true;
        }

        private async Task<bool> CreateMapOutputAsync(string mapLayerName, string mapLayerPath, string bufferLayerPath, string mapOutputType)
        {
            // Get the input feature class type.
            string mapLayerFCType = _mapFunctions.GetFeatureClassType(mapLayerName);
            if (mapLayerFCType == null)
                return false;

            // Get the buffer feature class type.
            string bufferFCType = _mapFunctions.GetFeatureClassType(_bufferLayerName);
            if (bufferFCType == null)
                return false;

            // If the input layer should be clipped to the buffer layer, do so now.
            if (mapOutputType == "CLIP")
            {
                if ((mapLayerFCType == "polygon" & bufferFCType == "polygon") ||
                    (mapLayerFCType == "line" & (bufferFCType == "line" || bufferFCType == "polygon")))
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
            else if (mapOutputType == "OVERLAY")
            {
                if ((bufferFCType == "polygon" & mapLayerFCType == "polygon") ||
                    (bufferFCType == "line" & (mapLayerFCType == "line" || mapLayerFCType == "polygon")))
                {
                    // Clip
                    FileFunctions.WriteLine(_logFile, "Overlaying selected features ...");
                    return await ArcGISFunctions.ClipFeaturesAsync(bufferLayerPath, mapLayerPath, _tempMasterOutputFile, true);
                }
                else
                {
                    // Select from the buffer layer.
                    FileFunctions.WriteLine(_logFile, "Selecting features  ...");
                    await ArcGISFunctions.SelectLayerByLocationAsync(bufferLayerPath, mapLayerPath);

                    // Find the buffer layer by name.
                    FeatureLayer bufferLayer = _mapFunctions.FindLayer(_bufferLayerName);

                    // Count the selected features.
                    int featureCount = bufferLayer.SelectionCount;
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

                    // Clear the buffer layer selection.
                    await _mapFunctions.ClearLayerSelectionAsync(_bufferLayerName);

                    return true;
                }
            }
            // If the input layer should be intersected with the buffer layer, do so now.
            else if (mapOutputType == "INTERSECT")
            {
                if ((mapLayerFCType == "polygon" & bufferFCType == "polygon") ||
                    (mapLayerFCType == "line" & bufferFCType == "line"))
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

        private async Task<bool> AddMapLabelsAsync(OverwriteLabelOptions overwriteLabelOption, bool overwriteLabels,
            String mapLabelColumn, string mapKeyColumn, string mapGroupName)
        {
            bool newLabelField = false;
            // Does the map label field already exist? If not, add it.
            if (!await _mapFunctions.FieldExistsAsync(_tempMasterLayerName, mapLabelColumn))
            {
                if (!await ArcGISFunctions.AddFieldAsync(_tempMasterOutputFile, mapLabelColumn, "LONG"))
                {
                    //MessageBox.Show("Error adding map label field '" + mapLabelColumn + "' to " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error adding map label field '" + mapLabelColumn + "' to " + _tempMasterOutputFile);

                    return false;
                }

                newLabelField = true;
            }

            // Either we have a new label field, or we want to overwrite the labels and are allowed to.
            if (newLabelField ||
                (overwriteLabelOption != OverwriteLabelOptions.No &&
                overwriteLabels))
            {
                // Add labels as required.
                if (!await CreateMapLabelsAsync(overwriteLabelOption, mapLabelColumn, mapKeyColumn, mapGroupName))
                {
                    //MessageBox.Show("Error setting map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error setting map labels to " + mapLabelColumn + " in " + _tempMasterOutputFile);

                    return false;
                }
            }

            return true;
        }

        private async Task<bool> CreateMapLabelsAsync(OverwriteLabelOptions overwriteLabelOption, String mapLabelColumn, string mapKeyColumn,
            string mapGroupName)
        {
            // Add relevant labels.
            if (overwriteLabelOption == OverwriteLabelOptions.ResetByLayer) // Reset each layer to 1.
            {
                FileFunctions.WriteLine(_logFile, "Resetting label counter ...");
                int startLabel = 1;
                FileFunctions.WriteLine(_logFile, "Adding map labels ...");

                if (await _mapFunctions.AddIncrementalNumbersAsync(_tempMasterOutputFile, _tempMasterLayerName, mapLabelColumn, mapKeyColumn, startLabel) < 0)
                {
                    MessageBox.Show("Error calculating map label field '" + mapLabelColumn + "' in " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error calculating map label field '" + mapLabelColumn + "' in " + _tempMasterOutputFile);

                    return false;
                }
            }
            else if (overwriteLabelOption == OverwriteLabelOptions.ResetByGroup && mapGroupName != "")
            {
                // Increment within but reset between groups. Note all group labels are already initialised as 1.
                // Only triggered if a group name has been found.

                int intGroupIndex = _mapGroupNames.IndexOf(mapGroupName);
                int intGroupLabel = _mapGroupLabels[intGroupIndex];
                FileFunctions.WriteLine(_logFile, "Adding map labels ...");

                intGroupLabel = await _mapFunctions.AddIncrementalNumbersAsync(_tempMasterOutputFile, _tempMasterLayerName, mapLabelColumn, mapKeyColumn, intGroupLabel);
                intGroupLabel++;
                _mapGroupLabels[intGroupIndex] = intGroupLabel; // Store the new max label.
            }
            else
            {
                // There is no group or groups are ignored, or we are not resetting. Use the existing max label number.
                int startLabel = _maxLabel;
                FileFunctions.WriteLine(_logFile, "Adding map labels ...");

                _maxLabel = await _mapFunctions.AddIncrementalNumbersAsync(_tempMasterOutputFile, _tempMasterLayerName, mapLabelColumn, mapKeyColumn, startLabel);
                _maxLabel++; // the new start label for incremental labeling
            }

            return true;
        }

        private async Task<int> ExportSelectionAsync(string outputTableName,
            string mapColumns, string mapGroupColumns, string mapStatsColumns, string mapOrderColumns,
            bool includeHeaders, string areaUnit,
            bool includeDistance, string radiusText)
        {
            int intLineCount = -999;

            // Only export if the user has specified columns.
            if (mapColumns == "")
                return -1;

            // Write out the results to table as appropriate.
            FileFunctions.WriteLine(_logFile, "Extracting summary information ...");

            // Check the input feature layer exists.
            FeatureLayer inputFeaturelayer = _mapFunctions.FindLayer(_tempMasterLayerName);
            if (inputFeaturelayer == null)
                return -1;

            // Get the input feature class type.
            string inputFeatureType = _mapFunctions.GetFeatureClassType(inputFeaturelayer);
            if (inputFeatureType == null)
                return -1;

            // Calculate the area field if required.
            if (!String.IsNullOrEmpty(areaUnit) && inputFeatureType == "polygon")
            {
                // Does the area field already exist? If not, add it.
                if (!await _mapFunctions.FieldExistsAsync(_tempMasterLayerName, "Area"))
                {
                    if (!await ArcGISFunctions.AddFieldAsync(_tempMasterOutputFile, "Area", "DOUBLE", 20))
                    {
                        //MessageBox.Show("Error adding area field to " + _tempMasterOutputFile + ".");
                        FileFunctions.WriteLine(_logFile, "Error adding area field to " + _tempMasterOutputFile);

                        return -1;
                    }
                }

                string geometryProperty = "Area AREA";
                if (areaUnit.Equals("ha", StringComparison.CurrentCultureIgnoreCase))
                {
                    areaUnit = "HECTARES";
                }
                else if (areaUnit.Equals("m2", StringComparison.CurrentCultureIgnoreCase))
                {
                    areaUnit = "SQUARE_METERS";
                }
                else if (areaUnit.Equals("km2", StringComparison.CurrentCultureIgnoreCase))
                {
                    areaUnit = "SQUARE_KILOMETERS";
                }

                // Calculate the area field.
                if (!await ArcGISFunctions.CalculateGeometryAsync(_tempMasterOutputFile, geometryProperty, "", areaUnit))
                {
                    //MessageBox.Show("Error calculating area field in " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error calculating area field in " + _tempMasterOutputFile);

                    return -1;
                }
            }

            // Calculate the distance if required.
            if (includeDistance)
            {
                // Now add the distance field by joining. This will take all fields.
                if (!await ArcGISFunctions.SpatialJoinAsync(_tempMasterOutputFile, _searchLayerName, _tempFCOutputFile,
                    "JOIN_ONE_TO_ONE", "KEEP_ALL", "", "CLOSEST", "0", "DISTANCE", addToMap: true))
                {
                    //MessageBox.Show("Error joining " + _tempMasterOutputFile  + " distance field to " + _tempMasterOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error joining " + _tempMasterOutputFile + " distance field to " + _tempMasterOutputFile);

                    return -1;
                }
            }
            else
            {
                // Do a straight copy of the input features.
                if (!await ArcGISFunctions.CopyFeaturesAsync(_tempMasterOutputFile, _tempFCOutputFile, true))
                {
                    //MessageBox.Show("Error copying output file to " + _tempFCOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error copying output file to " + _tempFCOutputFile);

                    return -1;
                }
            }

            // After this the input to the remainder of the function should be from _tempFCOutputFile.

            // Check the output feature layer exists.
            FeatureLayer outputFeatureLayer = _mapFunctions.FindLayer(_tempOutputLayerName);
            if (outputFeatureLayer == null)
                return -1;

            // Include radius if requested
            if (radiusText != "none")
            {
                FileFunctions.WriteLine(_logFile, "Including radius column ...");

                // Does the radius field already exist? If not, add it.
                if (!await _mapFunctions.FieldExistsAsync(_tempOutputLayerName, "Radius"))
                {
                    if (!await ArcGISFunctions.AddFieldAsync(_tempFCOutputFile, "Radius", "TEXT", fieldLength: 25))
                    {
                        //MessageBox.Show("Error adding radius field to " + _tempFCOutputFile + ".");
                        FileFunctions.WriteLine(_logFile, "Error adding radius field to " + _tempFCOutputFile);

                        return -1;
                    }
                }

                // Calculate the radius field.
                if (!await ArcGISFunctions.CalculateFieldAsync(_tempFCOutputFile, "Radius", '"' + radiusText + '"'))
                {
                    //MessageBox.Show("Error calculating radius field in " + _tempFCOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error calculating radius field in " + _tempFCOutputFile);

                    return -1;
                }
            }

            // Check all the requested group by and statistics fields exist.
            // Only pass those that do.
            if (mapGroupColumns != "")
            {
                List<string> groupColumnList = mapGroupColumns.Split(';').ToList();
                mapGroupColumns = "";
                foreach (string groupColumn in groupColumnList)
                {
                    string columnName = groupColumn.Trim();

                    if (!await _mapFunctions.FieldExistsAsync(_tempOutputLayerName, columnName))
                        mapGroupColumns = mapGroupColumns + columnName + ";";
                }
                if (mapGroupColumns != "")
                    mapGroupColumns = mapGroupColumns.Substring(0, mapGroupColumns.Length - 1);
            }

            if (mapStatsColumns != "")
            {
                List<string> statsColumnList = mapStatsColumns.Split(';').ToList();
                mapStatsColumns = "";
                foreach (string statsColumn in statsColumnList)
                {
                    List<string> statsComponents = statsColumn.Split(' ').ToList();
                    string columnName = statsComponents[0].Trim(); // The field name.

                    if (!await _mapFunctions.FieldExistsAsync(_tempOutputLayerName, columnName))
                        mapStatsColumns = mapStatsColumns + statsColumn + ";";
                }
                if (mapStatsColumns != "")
                    mapStatsColumns = mapStatsColumns.Substring(0, mapStatsColumns.Length - 1);
            }

            // If we have group columns but no statistics columns, add a dummy column.
            bool blDummyAdded = false;
            if (mapStatsColumns == "" && mapGroupColumns != "")
            {
                string strDummyField = mapGroupColumns.Split(';').ToList()[0];
                mapStatsColumns = strDummyField + " FIRST";
                blDummyAdded = true;
            }

            // Now do the summary statistics as required, or export the layer to table if not.
            if (mapGroupColumns != "" || mapStatsColumns != "")
            {
                FileFunctions.WriteLine(_logFile, "Calculating summary statistics ...");

                string statisticsFields = "";
                if (mapStatsColumns != "")
                    statisticsFields = mapStatsColumns;

                string caseFields = "";
                if (mapGroupColumns != "")
                    caseFields = mapGroupColumns;

                // Calculate the summary statistics.
                if (!await ArcGISFunctions.CalculateSummaryStatisticsAsync(_tempFCOutputFile, _tempTableOutputFile, statisticsFields, caseFields, addToMap: true))
                {
                    //MessageBox.Show("Error calculating radius field in " + _tempFCOutputFile + ".");
                    FileFunctions.WriteLine(_logFile, "Error calculating summary statistics for '" + _tempMasterOutputFile + "' into " + _tempFCOutputFile);

                    return -1;
                }

                //// Now rename the radius field.
                //if (radiusText != "none")
                //{
                //    ITable tpOutTable = GetTable(_tempOutputTableName, aLogFile, Messages);

                //    List<string> strFieldNames = mapStatsColumns.Split(';').ToList();
                //    int intIndexCount = 0;
                //    foreach (string strField in strFieldNames)
                //    {
                //        List<string> strFieldComponents = strField.Split(' ').ToList();
                //        string strNewField = strFieldComponents[0]; // The original name of the field.

                //        // Only rename the radius column
                //        if (strNewField.ToLower() == "radius")
                //        {
                //            // Let's find out what the new field is called - could be anything.
                //            int intNewIndex = 2; // OBJECTID = 0; Frequency = 1.
                //            intNewIndex = intNewIndex + mapGroupColumns.Split(';').ToList().Count + intIndexCount; // Add the number of columns used for grouping

                //            IField pNewField = tpOutTable.Fields.get_Field(intNewIndex);
                //            string strInputField = pNewField.Name;
                //            // Note index stays the same, since we're deleting the fields.

                //            // Get the definition of the original field from the original feature class.
                //            int intIndex = outputFeatureLayer.Fields.FindField(strNewField);
                //            IField pField = outputFeatureLayer.Fields.get_Field(intIndex);

                //            // Add the field to the new FC.
                //            AddTableField(_tempOutputTableName, strNewField, pField.Type, pField.Length, aLogFile, Messages);
                //            // Calculate the new field.
                //            string strCalc = "[" + strInputField + "]";
                //            CalculateField(tempTableFile, strNewField, strCalc, aLogFile, Messages);
                //            DeleteField(tpOutTable, strInputField, aLogFile, Messages);
                //        }
                //    }
                //}

                // Now export the output table.
                FileFunctions.WriteLine(_logFile, "Exporting output ...");
                intLineCount = await _mapFunctions.CopyTableToTextFileAsync(_tempOutputTableName, outputTableName, mapColumns, mapOrderColumns, false, true);
            }
            else
            {
                // Do straight copy of the feature class.
                FileFunctions.WriteLine(_logFile, "Exporting output ...");
                intLineCount = await _mapFunctions.CopyFCToTextFileAsync(_tempOutputLayerName, outputTableName, mapColumns, mapOrderColumns, false, true);
            }

            // Remove all temporary feature classes and tables.
            if (_mapFunctions.FindLayer(_tempOutputLayerName) != null)
                await _mapFunctions.RemoveLayerAsync(_tempOutputLayerName);

            if (_mapFunctions.FindTable(_tempOutputTableName) != null)
                await _mapFunctions.RemoveTableAsync(_tempOutputTableName);

            if (await ArcGISFunctions.FeatureClassExistsAsync(_tempFCOutputFile))
                ArcGISFunctions.DeleteFeatureClass(_tempGDBName, _tempOutputLayerName);

            if (await ArcGISFunctions.TableExistsAsync(_tempTableOutputFile))
                ArcGISFunctions.DeleteTable(_tempGDBName, _tempOutputTableName);

            return intLineCount;
        }

        private async Task<bool> KeepLayerAsync(string layerName, string outputFile, AddSelectedLayersOptions addSelectedLayersOption,
            string layerFileName, bool displayLabels, string labelClause, string labelColumn)
        {
            // Copy to a permanent file (note this is not the summarised layer).
            FileFunctions.WriteLine(_logFile, "Copying selected GIS features to " + outputFile + ".shp ...");
            await ArcGISFunctions.CopyFeaturesAsync(_tempMasterLayerName, outputFile);

            // If the layer is to be added to the map
            if (addSelectedLayersOption != AddSelectedLayersOptions.No)
            {
                // Set the layer symbology to use.
                string symbologyFile = _layerFolder + "\\" + layerFileName;

                if (!await SetLayerInMapAsync(layerName, symbologyFile))
                {
                    MessageBox.Show("Error setting output layer in the map.");
                    FileFunctions.WriteLine(_logFile, "Error setting output layer in the map");
                }

                FileFunctions.WriteLine(_logFile, "Output " + layerName + " added to display");

                // If labels are to be displayed
                if (addSelectedLayersOption == AddSelectedLayersOptions.WithLabels && displayLabels)
                {
                    // Translate the label string.
                    if (labelClause != "" && layerFileName == "") // Only if we don't have a layer file.
                    {
                        List<string> labelOptions = labelClause.Split('$').ToList();
                        string labelFont = labelOptions[0].Split(':')[1];
                        double labelSize = double.Parse(labelOptions[1].Split(':')[1]); // Needs error trapping
                        int labelRed = int.Parse(labelOptions[2].Split(':')[1]); // Needs error trapping
                        int labelGreen = int.Parse(labelOptions[3].Split(':')[1]);
                        int labelBlue = int.Parse(labelOptions[4].Split(':')[1]);
                        string labelOverlap = labelOptions[5].Split(':')[1];

                        await _mapFunctions.LabelLayerAsync(layerName, "[" + labelColumn + "]", labelFont, labelSize, "Normal",
                            labelRed, labelGreen, labelBlue, labelOverlap);
                    }
                    else if (labelColumn != "" && layerFileName == "")
                    {
                        await _mapFunctions.LabelLayerAsync(layerName, "[" + labelColumn + "]");
                    }

                    FileFunctions.WriteLine(_logFile, "Labels added to output " + layerName);
                }
                else
                {
                    // Turn labels off
                    await _mapFunctions.SwitchLabelsAsync(layerName, displayLabels);
                }
            }
            else
            {
                // User doesn't want to add the layer to the display.
                await _mapFunctions.RemoveLayerAsync(layerName);
            }

            return true;
        }

        #endregion Methods

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

        public string LayerGroup { get; set; }

        public string LayerName { get; set; }

        public string GISOutputName { get; set; }

        public string TableOutputName { get; set; }

        public string Columns { get; set; }

        public string GroupColumns { get; set; }

        public string StatisticsColumns { get; set; }

        public string OrderColumns { get; set; }

        public string Criteria { get; set; }

        public bool IncludeArea { get; set; }

        public bool IncludeDistance { get; set; }

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

        public bool IsOpen { get; set; }

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