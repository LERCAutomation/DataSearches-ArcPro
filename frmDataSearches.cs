
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;

using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.GeoDatabaseUI;

using HLSearchesToolConfig;
using HLFileFunctions;
using HLStringFunctions;
using HLArcMapModule;
using HLSearchesToolLaunchConfig;
using DataSearches.Properties;

using System.Data.OleDb;

namespace DataSearches
{
    public partial class frmDataSearches : Form
    {
        private void btnOK_Click(object sender, EventArgs e)
        {


















            }





            // All done, bring to front etc. 
            _mapFunctions.UpdateTOC();
            _mapFunctions.ToggleTOC();
            _mapFunctions.ToggleDrawing(true);
            _mapFunctions.SetContentsView();



        }


    }
}
