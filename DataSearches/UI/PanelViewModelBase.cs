﻿// The Data tools are a suite of ArcGIS Pro addins used to extract
// and manage biodiversity information from ArcGIS Pro and SQL Server
// based on pre-defined or user specified criteria.
//
// Copyright © 2024 Andy Foy Consulting.
//
// This file is part of DataSelector.
//
// DataSelector is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// DataSelector is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with DataSelector.  If not, see <http://www.gnu.org/licenses/>.

using ArcGIS.Desktop.Framework.Contracts;

namespace DataSearches.UI
{
    /// <summary>
    /// Abstract base class for panel view-models.
    /// </summary>
    internal abstract class PanelViewModelBase : PropertyChangedBase
    {
        #region Properties

        /// <summary>
        /// Panel display name.
        /// </summary>
        public abstract string DisplayName { get; }

        #endregion Properties
    }
}