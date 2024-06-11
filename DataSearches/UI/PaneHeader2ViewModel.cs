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
using Microsoft.Data.SqlClient;
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
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
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
                return ((!_dockPane.LayersListLoading)
                    //&& (_requireSiteName) ???
                    && (!_dockPane.SearchRunning));
            }
        }

        /// <summary>
        /// Is the list of layers enabled?
        /// </summary>
        public bool LayersListEnabled
        {
            get
            {
                return ((!_dockPane.LayersListLoading)
                    && (_layersList != null)
                    && (!_dockPane.SearchRunning));
            }
        }

        /// <summary>
        /// Is the list of buffer units options enabled?
        /// </summary>
        public bool BufferUnitsListEnabled
        {
            get
            {
                return ((!_dockPane.LayersListLoading)
                    && (_bufferUnitsList != null)
                    && (!_dockPane.SearchRunning));
            }
        }

        /// <summary>
        /// Is the list of add to map options enabled?
        /// </summary>
        public bool AddToMapListEnabled
        {
            get
            {
                return ((!_dockPane.LayersListLoading)
                    && (_addToMapList != null)
                    && (!_dockPane.SearchRunning));
            }
        }

        /// <summary>
        /// Is the list of overwrite options enabled?
        /// </summary>
        public bool OverwriteLabelsListEnabled
        {
            get
            {
                return ((!_dockPane.LayersListLoading)
                    && (_overwriteLabelsList != null)
                    && (!_dockPane.SearchRunning));
            }
        }

        /// <summary>
        /// Is the list of combined sites options enabled?
        /// </summary>
        public bool CombinedSitesListEnabled
        {
            get
            {
                return ((!_dockPane.LayersListLoading)
                    && (_combinedSitesList != null)
                    && (!_dockPane.SearchRunning));
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
                return ((!_dockPane.SearchRunning)
                    && (!_dockPane.LayersListLoading));
            }
        }
        /// <summary>
        /// Can the run button be pressed?
        /// </summary>
        public bool RunButtonEnabled
        {
            get
            {
                return ((_layersList != null)
                    && (!_dockPane.LayersListLoading)
                    && (!_dockPane.SearchRunning)
                    && (_selectedLayers != null) //???
                    && (_searchRefText != null)
                    //&& (!_requireSiteName || _siteNameText != null)   ???
                    && (_bufferSizeText != null)
                    && (_selectedBufferUnits != null)
                    && (_selectedAddToMap != null)
                    && (_selectedOverwriteLabels != null)
                    && (_selectedCombinedSites != null));
            }
        }

        #endregion Controls Enabled

        #region Controls Visibility

        /// <summary>
        /// Is the list of add to map options visible?
        /// </summary>
        public bool AddToMapListVisibility
        {
            get
            {
                return (_toolConfig.DefaultAddSelectedLayers == 0);
            }
        }

        /// <summary>
        /// Is the list of overwrite options visible?
        /// </summary>
        public bool OverwriteLabelsListVisibility
        {
            get
            {
                return (_toolConfig.DefaultOverwriteLabels == 0);
            }
        }

        /// <summary>
        /// Is the list of combined sites options visible?
        /// </summary>
        public bool CombinedSitesListVisibility
        {
            get
            {
                return (_toolConfig.DefaultCombinedSitesTable == 0);
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
            // Reset all of the fields.
            SearchRefText = null;
            //SiteNameText = null;  ???
            //SelectedLayers = null;
            BufferSizeText = _toolConfig.DefaultBufferSize.ToString();

            BufferUnitsList = _toolConfig.BufferUnitOptionsDisplay;
            if (_toolConfig.DefaultBufferUnit > 0)
                SelectedBufferUnits = _toolConfig.BufferUnitOptionsDisplay[_toolConfig.DefaultBufferUnit - 1];

            AddToMapList = _toolConfig.AddSelectedLayersOptions;
            if (_toolConfig.DefaultAddSelectedLayers > 0)
                SelectedAddToMap = _toolConfig.AddSelectedLayersOptions[_toolConfig.DefaultAddSelectedLayers - 1];

            OverwriteLabelsList = _toolConfig.OverwriteLabelOptions;
            if (_toolConfig.DefaultOverwriteLabels > 0)
                SelectedOverwriteLabels = _toolConfig.OverwriteLabelOptions[_toolConfig.DefaultOverwriteLabels - 1];

            CombinedSitesList = _toolConfig.CombinedSitesTableOptions;
            if (_toolConfig.DefaultCombinedSitesTable > 0)
                SelectedCombinedSites = _toolConfig.CombinedSitesTableOptions[_toolConfig.DefaultCombinedSitesTable - 1];

            ClearLogFile = _toolConfig.DefaultClearLogFile;
            OpenLogFile = _toolConfig.DefaultOpenLogFile;

            // Update the fields and buttons in the form.
            //PreSelectLayers();    ???
            OnPropertyChanged(nameof(BufferSizeText));
            OnPropertyChanged(nameof(BufferUnitsList));
            OnPropertyChanged(nameof(SelectedBufferUnits));
            OnPropertyChanged(nameof(AddToMapList));
            OnPropertyChanged(nameof(SelectedAddToMap));
            OnPropertyChanged(nameof(OverwriteLabelsList));
            OnPropertyChanged(nameof(SelectedOverwriteLabels));
            OnPropertyChanged(nameof(CombinedSitesList));
            OnPropertyChanged(nameof(SelectedCombinedSites));
            OnPropertyChanged(nameof(RunButtonEnabled));
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
            // Replace any illegal characters in the user name string.
            string userID = StringFunctions.StripIllegals(Environment.UserName, "_", false);

            // User ID should be something at least.
            if (string.IsNullOrEmpty(userID))
            {
                userID = "Temp";
                FileFunctions.WriteLine(_logFile, "User ID not found. User ID used will be 'Temp'");
            }






            //// Set the destination log file path.
            //_logFile = _toolConfig.GetLogFilePath + @"\DataSearches_" + userID + ".log";

            //// Clear the log file if required.
            //if (ClearLogFile)
            //{
            //    bool blDeleted = FileFunctions.DeleteFile(_logFile);
            //    if (!blDeleted)
            //    {
            //        MessageBox.Show("Cannot delete log file. Please make sure it is not open in another window.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Error);
            //        return;
            //    }
            //}

            //// Check the user entered parameters.
            //if (string.IsNullOrEmpty(ColumnsText))
            //{
            //    MessageBox.Show("Please specify which columns you wish to select.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
            //    return;
            //}

            //// Table name should always be selected.
            //if (string.IsNullOrEmpty(SelectedTable)
            //        && (WhereText.Length <= 5 || !WhereText.Substring(0, 5).Equals("from ", StringComparison.CurrentCultureIgnoreCase)))
            //{
            //    MessageBox.Show("Please select a table to query from.", "Data Searches.", MessageBoxButton.OK, MessageBoxImage.Information);
            //    return;
            //}

            //// Output format should always be selected.
            //if (string.IsNullOrEmpty(SelectedOutputFormat))
            //{
            //    MessageBox.Show("Please select an output format.", "Data Searches", MessageBoxButton.OK, MessageBoxImage.Information);
            //    return;
            //}

            //// Update the fields and buttons in the form.
            //OnPropertyChanged(nameof(TablesList));
            //OnPropertyChanged(nameof(TablesListEnabled));
            //OnPropertyChanged(nameof(LoadColumnsEnabled));
            //OnPropertyChanged(nameof(ClearButtonEnabled));
            //OnPropertyChanged(nameof(SaveButtonEnabled));
            //OnPropertyChanged(nameof(LoadButtonEnabled));
            //OnPropertyChanged(nameof(VerifyButtonEnabled));
            //OnPropertyChanged(nameof(RunButtonEnabled));
            //_dockPane.RefreshPanel1Buttons();

            //// Perform the selection.
            //await ExecuteSelectionAsync(userID); // Success not currently checked.

            //// Update the fields and buttons in the form.
            //OnPropertyChanged(nameof(TablesList));
            //OnPropertyChanged(nameof(TablesListEnabled));
            //OnPropertyChanged(nameof(LoadColumnsEnabled));
            //OnPropertyChanged(nameof(ClearButtonEnabled));
            //OnPropertyChanged(nameof(SaveButtonEnabled));
            //OnPropertyChanged(nameof(LoadButtonEnabled));
            //OnPropertyChanged(nameof(VerifyButtonEnabled));
            //OnPropertyChanged(nameof(RunButtonEnabled));
            //_dockPane.RefreshPanel1Buttons();
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
                //OnPropertyChanged(nameof(SaveButtonEnabled)); ???
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
                //OnPropertyChanged(nameof(SaveButtonEnabled)); ???
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

        private ObservableCollection<string> _layersList;

        /// <summary>
        /// Get the list of loaded GIS layers.
        /// </summary>
        public ObservableCollection<string> LayersList
        {
            get
            {
                return _layersList;
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
                //OnPropertyChanged(nameof(RunButtonEnabled));  ???
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
                //OnPropertyChanged(nameof(SaveButtonEnabled)); ???
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
                //OnPropertyChanged(nameof(BufferUnitsListEnabled));
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
                //OnPropertyChanged(nameof(SaveButtonEnabled)); ???
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
                //OnPropertyChanged(nameof(SaveButtonEnabled)); ???
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
                //OnPropertyChanged(nameof(SaveButtonEnabled)); ???
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
        /// Load the list of open GIS layers.
        /// </summary>
        /// <param name="selectedTable"></param>
        /// <returns></returns>
        public async Task LoadLayersAsync(List<string> preSelectLayers)
        {
            //if (!string.IsNullOrEmpty(_columnsText))
            //{
            //    MessageBoxResult dlResult = MessageBox.Show("There is already text in the Columns field. Do you want to overwrite it?", "Data Searches", MessageBoxButton.YesNo, MessageBoxImage.Question);
            //    if (dlResult == MessageBoxResult.No)
            //        return; //User clicked by accident; leave routine.
            //}

            //// Get the field names and wait for the task to finish.
            //List<string> columnsList = await _sqlFunctions.GetFieldNamesListAsync(SelectedTable);

            //// Convert the field names to a single string.
            //string columnNamesText = "";
            //foreach (string columnName in columnsList)
            //{
            //    columnNamesText = columnNamesText + columnName + ",\r\n";
            //}
            //columnNamesText = columnNamesText.Substring(0, columnNamesText.Length - 3);

            //// Replace the text box value with the field names.
            //_columnsText = columnNamesText;

            //// Update the fields and buttons in the form.
            //OnPropertyChanged(nameof(ColumnsText));
            //OnPropertyChanged(nameof(SaveButtonEnabled));
            //OnPropertyChanged(nameof(ClearButtonEnabled));
            //OnPropertyChanged(nameof(VerifyButtonEnabled));
            //OnPropertyChanged(nameof(RunButtonEnabled));
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
}