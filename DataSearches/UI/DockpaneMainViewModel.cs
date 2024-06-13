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

using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Controls;
using ArcGIS.Desktop.Mapping.Events;
using ArcGIS.Desktop.Mapping;
using DataSearches;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using MessageBox = ArcGIS.Desktop.Framework.Dialogs.MessageBox;
using System.ComponentModel;

namespace DataSearches.UI
{
    /// <summary>
    /// Build the DockPane.
    /// </summary>
    internal class DockpaneMainViewModel : DockPane, INotifyPropertyChanged
    {
        #region Fields

        private DockpaneMainViewModel _dockPane;

        private PaneHeader1ViewModel _paneH1VM;
        private PaneHeader2ViewModel _paneH2VM;

        private bool _subscribed;

        #endregion Fields

        #region ViewModelBase Members

        /// <summary>
        /// Set the global variables.
        /// </summary>
        protected DockpaneMainViewModel()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Initialise the DockPane components.
        /// </summary>
        public async void InitializeComponent()
        {
            _dockPane = this;
            _initialised = false;
            _inError = false;

            // Setup the tab controls.
            PrimaryMenuList.Clear();

            PrimaryMenuList.Add(new TabControl() { Text = "Profile", Tooltip = "Select XML profile" });
            PrimaryMenuList.Add(new TabControl() { Text = "Search", Tooltip = "Run data search" });

            // Load the default XML profile (or let the user choose a profile.
            _paneH1VM = new PaneHeader1ViewModel(_dockPane);

            // If the profile was in error.
            if (_paneH1VM.XMLError)
            {
                _inError = true;
                return;
            }

            // If the default (and only) profile was loaded.
            if (_paneH1VM.XMLLoaded)
            {
                // Initialise the search pane.
                bool initialised = await InitialiseSearchPaneAsync(false);
                if (!initialised)
                    return;

                // Select the profile tab.
                SelectedPanelHeaderIndex = 1;
            }
            else
            {
                // Select the search tab.
                SelectedPanelHeaderIndex = 0;
            }

            // Indicate that the dockpane has been initialised.
            _initialised = true;
        }

        /// <summary>
        /// Show the DockPane.
        /// </summary>
        internal static void Show()
        {
            // Get the dockpane DAML id.
            DockPane pane = FrameworkApplication.DockPaneManager.Find(_dockPaneID);
            if (pane == null)
                return;

            // Get the ViewModel by casting the dockpane.
            DockpaneMainViewModel vm = pane as DockpaneMainViewModel;

            // If the ViewModel is uninitialised then initialise it.
            if (!vm.Initialised)
                vm.InitializeComponent();

            // If the ViewModel is in error then don't show the dockpane.
            if (vm.InError)
            {
                pane = null;
                return;
            }

            // Active the dockpane.
            pane.Activate();
        }

        protected override void OnShow(bool isVisible)
        {
            // Is the dockpane visible (or is the window not showing the map).
            if (isVisible)
            {
                if (!_subscribed)
                {
                    _subscribed = true;

                    // Connect to map events.
                    ArcGIS.Desktop.Mapping.Events.ActiveMapViewChangedEvent.Subscribe(OnActiveMapViewChanged);
                    ArcGIS.Desktop.Mapping.Events.DrawCompleteEvent.Subscribe(OnDrawComplete);
                }
            }
            else
            {
                if (_subscribed)
                {
                    _subscribed = false;

                    // Unsubscribe from map events.
                    ArcGIS.Desktop.Mapping.Events.ActiveMapViewChangedEvent.Unsubscribe(OnActiveMapViewChanged);
                    ArcGIS.Desktop.Mapping.Events.DrawCompleteEvent.Unsubscribe(OnDrawComplete);
                }
            }
            base.OnShow(isVisible);
        }

        #endregion ViewModelBase Members

        #region Properties

        /// <summary>
        /// ID of the DockPane.
        /// </summary>
        private const string _dockPaneID = "DataSearches_UI_DockpaneMain";

        public static string DockPaneID
        {
            get => _dockPaneID;
        }

        /// <summary>
        /// Override the default behavior when the dockpane's help icon is clicked
        /// or the F1 key is pressed.
        /// </summary>
        protected override void OnHelpRequested()
        {
            if (_helpURL != null)
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = _helpURL,
                    UseShellExecute = true
                });
            }
        }

        private List<TabControl> _primaryMenuList = [];

        /// <summary>
        /// Get the list of dock panes.
        /// </summary>
        public List<TabControl> PrimaryMenuList
        {
            get { return _primaryMenuList; }
        }

        private int _selectedPanelHeaderIndex = 0;

        /// <summary>
        /// Get/Set the active pane.
        /// </summary>
        public int SelectedPanelHeaderIndex
        {
            get { return _selectedPanelHeaderIndex; }
            set
            {
                _selectedPanelHeaderIndex = value;
                OnPropertyChanged(nameof(SelectedPanelHeaderIndex));

                if (_selectedPanelHeaderIndex == 0)
                    CurrentPage = _paneH1VM;
                if (_selectedPanelHeaderIndex == 1)
                    CurrentPage = _paneH2VM;
            }
        }

        private PanelViewModelBase _currentPage;

        /// <summary>
        /// The currently active DockPane.
        /// </summary>
        public PanelViewModelBase CurrentPage
        {
            get { return _currentPage; }
            set
            {
                _currentPage = value;
                OnPropertyChanged(nameof(CurrentPage));
            }
        }

        private bool _initialised = false;

        /// <summary>
        /// Has the DockPane been initialised?
        /// </summary>
        public bool Initialised
        {
            get { return _initialised; }
            set
            {
                _initialised = value;
            }
        }

        private bool _inError = false;

        /// <summary>
        /// Is the DockPane in error?
        /// </summary>
        public bool InError
        {
            get { return _inError; }
            set
            {
                _inError = value;
            }
        }

        private bool _layersListLoading;

        /// <summary>
        /// Is the layers list loading?
        /// </summary>
        public bool LayersListLoading
        {
            get { return _layersListLoading; }
            set { _layersListLoading = value; }
        }

        private bool _searchRunning;

        /// <summary>
        /// Is the search running?
        /// </summary>
        public bool SearchRunning
        {
            get { return _searchRunning; }
            set { _searchRunning = value; }
        }

        private string _helpURL;

        /// <summary>
        /// The URL of the help page.
        /// </summary>
        public string HelpURL
        {
            get { return _helpURL; }
            set { _helpURL = value; }
        }

        #endregion Properties

        #region Methods

        private void OnActiveMapViewChanged(ActiveMapViewChangedEventArgs obj)
        {
            if (MapView.Active == null)
                DockpaneVisibility = Visibility.Hidden;
            else
            {
                DockpaneVisibility = Visibility.Visible;

                // Reload the form layers (don't wait for the response).
                _paneH2VM.LoadLayersAsync(false);
            }

        }

        private void OnDrawComplete(MapViewEventArgs obj)
        {
            if (MapView.Active == null)
                DockpaneVisibility = Visibility.Hidden;
            else
            {
                DockpaneVisibility = Visibility.Visible;

                // Reload the form layers (don't wait for the response).
                _paneH2VM.LoadLayersAsync(false);
            }
        }

        private Visibility _dockpaneVisibility = Visibility.Hidden;
        public Visibility DockpaneVisibility
        {
            get { return _dockpaneVisibility; }
            set
            {
                _dockpaneVisibility = value;
                OnPropertyChanged(nameof(DockpaneVisibility));
            }
        }

        /// <summary>
        /// Initialise the search pane.
        /// </summary>
        /// <returns></returns>
        public async Task<bool> InitialiseSearchPaneAsync(bool messages)
        {
            _paneH2VM = new PaneHeader2ViewModel(_dockPane, _paneH1VM.ToolConfig);

            // Load the form (don't wait for the response).
            Task.Run(() => _paneH2VM.ResetForm(false));

            return true;
        }

        /// <summary>
        /// Reset the search pane.
        /// </summary>
        public void ClearSearchPane()
        {
            _paneH2VM = null;
        }

        /// <summary>
        /// Event when the DockPane is hidden.
        /// </summary>
        protected override void OnHidden()
        {
            // Get the dockpane DAML id.
            DockPane pane = FrameworkApplication.DockPaneManager.Find(_dockPaneID);
            if (pane == null)
                return;

            // Get the ViewModel by casting the dockpane.
            DockpaneMainViewModel vm = pane as DockpaneMainViewModel;

            // Force the dockpane to be re-initialised next time it's shown.
            vm.Initialised = false;
        }

        public void RefreshPanel1Buttons()
        {
            _paneH1VM.RefreshButtons();
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
    /// Button implementation to show the DockPane.
    /// </summary>
    internal class DockpaneMain_ShowButton : Button
    {
        protected override void OnClick()
        {
            //string uri = System.Reflection.Assembly.GetExecutingAssembly().Location;

            // Show the dock pane.
            DockpaneMainViewModel.Show();
        }
    }
}