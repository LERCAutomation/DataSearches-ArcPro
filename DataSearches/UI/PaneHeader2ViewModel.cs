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

// Ignore Spelling: Symbology Tooltip Img

using ArcGIS.Core.Data;
using ArcGIS.Desktop.Framework;
using DataSearches;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using DataSearches.UI;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;

namespace DataSearches.UI
{
    internal class PaneHeader2ViewModel : PanelViewModelBase, INotifyPropertyChanged
    {
        #region Fields

        private readonly DockpaneMainViewModel _dockPane;

        private string _logFile;

        private const string _displayName = "DataSearches";

        private readonly DataSearchesConfig _toolConfig;

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
            //// Get the relevant config file settings.
            //_includeWildcard = _toolConfig.GetIncludeWildcard;
            //_excludeWildcard = _toolConfig.GetExcludeWildcard;
            //_defaultFormat = _toolConfig.GetDefaultFormat;
            //_defaultSchema = _toolConfig.GetDatabaseSchema;
            //_clearLogFile = _toolConfig.GetDefaultClearLogFile;
            //_openLogFile = _toolConfig.GetDefaultOpenLogFile;
            //_setSymbology = _toolConfig.GetDefaultSetSymbology;
            ////_layerLocation = _toolConfig.GetLayerLocation;
            //_validateSQL = _toolConfig.GetValidateSQL;

            //// Set the window properties.
            //_selectedOutputFormat = _defaultFormat;
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
                    && (!_toolConfig.RequireSiteName || _siteNameText != null)
                    && (!String.IsNullOrEmpty(_bufferSizeText))
                    && (_selectedBufferUnits != null)
                    && (_toolConfig.DefaultAddSelectedLayers <= 0 || _selectedAddToMap != null)
                    && (_toolConfig.DefaultOverwriteLabels <= 0 || _selectedOverwriteLabels != null)
                    && (_toolConfig.DefaultCombinedSitesTable <= 0 || _selectedCombinedSites != null));
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
                if (_toolConfig.DefaultAddSelectedLayers == -1)
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
                if (_toolConfig.DefaultAddSelectedLayers == -1)
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
                if (_toolConfig.DefaultCombinedSitesTable == -1)
                    return Visibility.Collapsed;
                else
                    return Visibility.Visible;
            }
        }

        #endregion Controls Visibility

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
            if (_toolConfig.RequireSiteName && string.IsNullOrEmpty(SiteNameText))
            {
                MessageBox.Show("Please enter a location.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!OpenLayersList.Where(p => p.IsSelected).Any())
            {
                MessageBox.Show("Please select at least one layer to search.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (string.IsNullOrEmpty(BufferSizeText))
            {
                MessageBox.Show("Please enter a buffer size.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            bool bufferNumeric = Double.TryParse(BufferSizeText, out double bufferNumber);
            if (!bufferNumeric || bufferNumber < 0) // User either entered text or a negative number
            {
                MessageBox.Show("Please enter a positive number for the buffer size.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (string.IsNullOrEmpty(SelectedBufferUnits))
            {
                MessageBox.Show("Please enter a buffer size.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
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

        private ObservableCollection<Layers> _openLayersList;

        /// <summary>
        /// Get the list of loaded GIS layers.
        /// </summary>
        public ObservableCollection<Layers> OpenLayersList
        {
            get
            {
                return _openLayersList;
            }
        }

        private List<string> _selectedLayers;

        /// <summary>
        /// Get/Set the selected loaded layers.
        /// </summary>
        public List<string> SelectedLayers
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

        private string _selectedBufferUnits;

        /// <summary>
        /// Get/Set the select buffer size units.
        /// </summary>
        public string SelectedBufferUnits
        {
            get
            {
                return _selectedBufferUnits;
            }
            set
            {
                _selectedBufferUnits = value;

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

        public Visibility ProcessingAnimation
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
            OnPropertyChanged(nameof(SelectedBufferUnits));
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
                foreach (Layers layer in _openLayersList)
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
                SelectedBufferUnits = _toolConfig.BufferUnitOptionsDisplay[_toolConfig.DefaultBufferUnit - 1];

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
            await LoadLayersAsync(reset);
        }

        /// <summary>
        /// Load the list of open GIS layers.
        /// </summary>
        /// <param name="selectedTable"></param>
        /// <returns></returns>
        public async Task LoadLayersAsync(bool reset)
        {
            _dockPane.LayersListLoading = true;

            if (reset)
                _processingLabel = "Resetting form ...";
            else
                _processingLabel = "Loading form ...";
            OnPropertyChanged(nameof(ProcessingLabel));
            OnPropertyChanged(nameof(ProcessingAnimation));

            // Update the fields and buttons in the form.
            UpdateFormControls();

            // Temporary sleep to check processing label and animation are working. ???
            System.Threading.Thread.Sleep(100);

            List<string> closedLayers = []; // The closed layers by name.

            await Task.Run(() =>
            {
                // Create a new map functions object.
                MapFunctions mapFunctions = new();

                // Check if there is an active map.
                bool mapOpen = (mapFunctions.MapName != null);

                // Reset the list of open layers.
                ObservableCollection<Layers> openLayersList = [];

                if (mapOpen)
                {
                    List<Layers> allLayers = _toolConfig.MapLayers;
                    //List<string> AllDisplayLayers = _toolConfig.MapNames; // All possible layers by display name
                    //List<bool> blLoadWarnings = _toolConfig.MapLoadWarnings; // A list telling us whether to warn users if layer not present
                    //List<bool> blPreselectLayers = _toolConfig.MapPreselectLayers; // A list telling us which layers to preselect in the form
                    //ObservableCollection<string> OpenLayers = []; // The open layers by name
                    //List<bool> PreselectLayers = []; // The preselect options of the open layers

                    // Loop through all of the layers to check if they are open
                    // in the active map.
                    foreach (Layers layer in allLayers)
                    {
                        if (mapFunctions.FindLayer(layer.LayerName) != null)
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

            _processingLabel = null;
            OnPropertyChanged(nameof(ProcessingLabel));
            OnPropertyChanged(nameof(ProcessingAnimation));

            _dockPane.LayersListLoading = false;

            // Update the fields and buttons in the form.
            UpdateFormControls ();

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
            // Save the parameters.
            string searchRef = SearchRefText;
            string siteName = SiteNameText;
            //string strOrganisation = txtOrganisation.Text; // Associated organisation




            bool blUpdateTable = _toolConfig.UpdateTable;

            // Replace any illegal characters in the user name string.
            string userID = StringFunctions.StripIllegals(Environment.UserName, "_", false);

            // User ID should be something at least.
            if (string.IsNullOrEmpty(userID))
            {
                userID = "Temp";
                FileFunctions.WriteLine(_logFile, "User ID not found. User ID used will be 'Temp'");
            }

            string strSaveRootDir = _toolConfig.SaveRootDir;    //??? remove str prefix


            // Indicate the search has started.
            _dockPane.SearchRunning = true;
            _processingLabel = "Performing search ...";
            OnPropertyChanged(nameof(ProcessingLabel));
            OnPropertyChanged(nameof(ProcessingAnimation));

            // Temporary delay to check processing label and animation are working. ???
            await Task.Delay(1000);




            FileFunctions.WriteLine(_logFile, "---------------------------------------------------------------------------");
            FileFunctions.WriteLine(_logFile, "Process complete");
            FileFunctions.WriteLine(_logFile, "---------------------------------------------------------------------------");

            // Indicate search has finished.
            _dockPane.SearchRunning = false;
            _processingLabel = null;
            OnPropertyChanged(nameof(ProcessingLabel));
            OnPropertyChanged(nameof(ProcessingAnimation));

            MessageBox.Show(String.Format("Search '{0}' complete!", searchRef), "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);

            // Open the log file (if required).
            if (OpenLogFile)
                Process.Start("notepad.exe", _logFile);

            return;
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
    /// GIS layers to search.
    /// </summary>
    public class Layers : INotifyPropertyChanged
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

        public Layers()
        {
        }

        public Layers(string nodeName)
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