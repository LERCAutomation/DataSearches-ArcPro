
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








            // Update the table if required.
            if (updateTable && (!string.IsNullOrEmpty(siteColumn) || !string.IsNullOrEmpty(orgColumn) || !string.IsNullOrEmpty(radiusColumn)))
            {
                FileFunctions.WriteLine(_logFile, "Updating attributes in search layer ...");

                if (!string.IsNullOrEmpty(siteColumn) && !string.IsNullOrEmpty(siteName))
                {
                    //FileFunctions.WriteLine(_logFile, "Updating " + siteColumn + " with '" + siteName + "'");
                    MapFunctions.CalculateField(targetLayer, siteColumn, '"' + siteName + '"', _logFile);
                }

                if (!string.IsNullOrEmpty(orgColumn) && !string.IsNullOrEmpty(strOrganisation))
                {
                    MapFunctions.CalculateField(targetLayer, orgColumn, '"' + strOrganisation + '"', _logFile);
                }

                if (!string.IsNullOrEmpty(radiusColumn) && !string.IsNullOrEmpty(radius))
                {
                    MapFunctions.CalculateField(targetLayer, radiusColumn, '"' + radius + '"', _logFile);
                }
            }











                    // Add to combined sites table as appropriate
                    // Function to take account of group by, order by and statistics fields.
                    if (mapCombinedSitesColumns != "" && combinedTableCreate)
                    {
                        FileFunctions.WriteLine(_logFile, "Extracting summary output for combined sites table ...");

                        intLineCount = _mapFunctions.ExportSelectionToCSV(tempMasterLayerName, strCombinedTable, mapCombinedSitesColumns, false, tempFCOutput, tempTableOutput, mapCombinedSitesGroupColumns,
                            mapCombinedSitesStatsColumns, mapCombinedSitesOrderColumns, includeArea, areaMeasureUnit, includeDistance, radiusText, targetLayer, _logFile, false, RenameColumns: true);

                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " row(s) added to summary output");
                    }

                }


                // Trigger the macro if one exists
                if (mapMacroName != "") // Only if we have a macro to trigger.
                {
                    FileFunctions.WriteLine(_logFile, "Executing vbscript macro ...");

                    Process scriptProc = new Process();
                    //scriptProc.StartInfo.FileName = @"C:\Windows\SysWOW64\cscript.exe";
                    scriptProc.StartInfo.FileName = @"cscript.exe";
                    scriptProc.StartInfo.WorkingDirectory = FileFunctions.GetDirectoryName(mapMacroName); //<---very important
                    scriptProc.StartInfo.UseShellExecute = true;
                    scriptProc.StartInfo.Arguments = string.Format(@"//B //Nologo {0} {1} {2} {3}", "\"" + mapMacroName + "\"", "\"" + _outputFolder + "\"", "\"" + mapTableOutputName + "." + mapFormat.ToLower() + "\"", "\"" + mapTableOutputName + ".xlsx" + "\"");
                    scriptProc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden; //prevent console window from popping up

                    try
                    {
                        scriptProc.Start();
                        scriptProc.WaitForExit(); // <-- Optional if you want program running until your script exit

                        int exitcode = scriptProc.ExitCode;
                        if (exitcode != 0)
                            FileFunctions.WriteLine(_logFile, "Error executing vbscript macro. Exit code : " + exitcode);

                        scriptProc.Close();
                    }
                    catch
                    {
                        MessageBox.Show("Error executing vbscript macro.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }

            }





            // All done, bring to front etc. 
            _mapFunctions.UpdateTOC();
            _mapFunctions.ToggleTOC();
            _mapFunctions.ToggleDrawing(true);
            _mapFunctions.SetContentsView();
            if (bufferSize != "0" && blKeepBuffer)
                _mapFunctions.ZoomToLayer(strLayerName, _logFile);



        }




        private void txtSearch_TextChanged(object sender, EventArgs e)
        {

            if (txtSearch.Text != "" && txtSearch.Text.Length >= 2) // Only fire it when it looks like we have a complete reference.
            {
                // Do we have a database name? If so, look up the reference.
                if (_toolConfig.Database != "")
                {
                    //-------------------------------------------------------------
                    // Use connection string for .accdb or .mdb as appropriate
                    //-------------------------------------------------------------
                    string strAccessConn;
                    if (FileFunctions.GetExtension(_toolConfig.Database).ToLower() == "accdb")
                    {
                        strAccessConn = "Provider='Microsoft.ACE.OLEDB.12.0';data source='" + _toolConfig.Database + "'";
                    }
                    else
                    {
                        strAccessConn = "Provider='Microsoft.Jet.OLEDB.4.0';data source='" + _toolConfig.Database + "'";
                    }

                    string siteColumn = _toolConfig.SiteColumn;
                    string orgColumn = _toolConfig.OrgColumn;
                    string strTable = _toolConfig.Table;
                    if (string.IsNullOrEmpty(strTable))
                        strTable = "Enquiries";

                    string mapColumns = siteColumn;
                    if (!string.IsNullOrEmpty(orgColumn))
                    {
                        if (!string.IsNullOrEmpty(mapColumns))
                            mapColumns = mapColumns + "," + orgColumn;
                        else
                            mapColumns = orgColumn;
                    }

                    string searchQuery = "SELECT " + mapColumns + " from " + strTable + " WHERE LCASE(" + _toolConfig.RefColumn + ") = " + '"' + txtSearch.Text.ToLower() + '"';
                    string strLocation = "";
                    string strOrganisation = "";

                    OleDbConnection myAccessConn = null;
                    try
                    {
                        myAccessConn = new OleDbConnection(strAccessConn);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Failed to create a database connection. System error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    DataSet myDataSet = new DataSet();
                    try
                    {

                        OleDbCommand myAccessCommand = new OleDbCommand(searchQuery, myAccessConn);
                        OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                        myAccessConn.Open();
                        myDataAdapter.Fill(myDataSet, strTable);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Failed to retrieve the required data from the database. System error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    finally
                    {
                        myAccessConn.Close();
                    }

                    DataRowCollection myRS = myDataSet.Tables[strTable].Rows;
                    foreach (DataRow aRow in myRS) // Really there should only be one. We can check for this.
                    {
                        // Get the location and organisation names
                        strLocation = aRow[0].ToString();
                        if (!string.IsNullOrEmpty(orgColumn))
                            strOrganisation = aRow[1].ToString();
                    }

                    if (strLocation != "")
                    {
                        // The location is known. Fill it in and do not allow editing.
                        txtLocation.Text = strLocation;
                        txtLocation.Enabled = false;
                    }
                    else
                    {
                        // The location is not known. Allow user to enter it.
                        txtLocation.Text = "";
                        txtLocation.Enabled = true;
                        // Should we allow an update??
                    }

                    txtOrganisation.Text = strOrganisation;

                }
                else
                {
                    txtLocation.Enabled = true;
                }
            }
            else
            {
                // There is no search reference. Disable the Location text box.
                txtLocation.Text = "";
                txtLocation.Enabled = false;
                txtOrganisation.Text = "";
            }
        }
    }
}
