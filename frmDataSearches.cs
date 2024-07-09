
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


            // The selected layers
            List<string> SelectedLayers = new List<string>();
            foreach (string strSelectedItem in lstLayers.SelectedItems)
            {
                SelectedLayers.Add(strSelectedItem);
            }












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











            // Now go through the layers.

            // Get any groups and initialise required layers.
            List<string> liGroupNames = new List<string>();
            List<int> liGroupLabels = new List<int>();
            if (resetGroups)
            {
                liGroupNames = StringFunctions.ExtractGroups(SelectedLayers);
                foreach (string strGroupName in liGroupNames)
                {
                    liGroupLabels.Add(1); // each group has its own label counter.
                }
            }






                    //string strTempShapeFile = tempFolder + @"\TempShapes_" + strUserID + ".shp";
                    //string strTempShapeOutput = tempFolder + @"\" + strTempOutput + ".shp";
                    //string strTempDBFOutput = tempFolder + @"\" + strTempOutput + "DBF.dbf";
                    string radiusText = "none";
                    if (includeRadius) radiusText = radius; // Only include radius if requested.

                    // Only export if the user has specified columns.
                    int intLineCount = -999;
                    if (mapColumns != "")
                    {
                        // Write out the results to table as appropriate.
                        bool blIncHeaders = false;
                        if (mapFormat.ToLower() == "csv") blIncHeaders = true;

                        FileFunctions.WriteLine(_logFile, "Extracting summary information ...");

                        intLineCount = _mapFunctions.ExportSelectionToCSV(tempMasterLayerName, outputTableName, mapColumns, blIncHeaders, tempFCOutput, tempTableOutput,
                            mapGroupColumns, mapStatsColumns, mapOrderColumns, includeArea, areaMeasureUnit, includeDistance, radiusText, targetLayer, _logFile, RenameColumns: true);

                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " record(s) exported");
                    }

                    // Copy to permanent layer as appropriate
                    if (keepLayer)
                    {
                        // Keep the layer - write to permanent file (note this is not the summarised layer).
                        FileFunctions.WriteLine(_logFile, "Copying selected GIS features to " + strShapeLayerName + ".shp ...");
                        _mapFunctions.CopyFeatures(tempMasterLayerName, outputFileName, aLogFile: _logFile);

                        // If the layer is to be added to the map
                        if (addSelectedLayers.ToLower() != "no")
                        {
                            // Add the permanent layer to the map (by moving it to the group layer)
                            _mapFunctions.MoveToGroupLayer(groupLayerName, _mapFunctions.GetLayer(strShapeLayerName, _logFile), _logFile);

                            // If there is a layer file
                            if (mapLayerFileName != "")
                            {
                                string strDisplayLayerFile = _layerFolder + @"\" + mapLayerFileName;
                                _mapFunctions.ChangeLegend(strShapeLayerName, strDisplayLayerFile, displayLabels, _logFile);
                            }

                            FileFunctions.WriteLine(_logFile, "Output " + strShapeLayerName + " added to display");

                            // If labels are to be displayed
                            if (addSelectedLayers.ToLower().Contains("with ") && displayLabels)
                            {
                                // Translate the label string.
                                if (mapLabelClause != "" && mapLayerFileName == "") // Only if we don't have a layer file.
                                {
                                    List<string> strLabelOptions = mapLabelClause.Split('$').ToList();
                                    string strFont = strLabelOptions[0].Split(':')[1];
                                    double dblSize = double.Parse(strLabelOptions[1].Split(':')[1]); // Needs error trapping
                                    int intRed = int.Parse(strLabelOptions[2].Split(':')[1]); // Needs error trapping
                                    int intGreen = int.Parse(strLabelOptions[3].Split(':')[1]);
                                    int intBlue = int.Parse(strLabelOptions[4].Split(':')[1]);
                                    string strOverlap = strLabelOptions[5].Split(':')[1];
                                    _mapFunctions.AnnotateLayer(strShapeLayerName, "[" + mapLabelColumn + "]", strFont, dblSize,
                                        intRed, intGreen, intBlue, strOverlap, aLogFile: _logFile);
                                    FileFunctions.WriteLine(_logFile, "Labels added to output " + mapOutputName);
                                }
                                else if (mapLabelColumn != "" && mapLayerFileName == "")
                                {
                                    _mapFunctions.AnnotateLayer(mapOutputName, "[" + mapLabelColumn + "]", aLogFile: _logFile);
                                    FileFunctions.WriteLine(_logFile, "Labels added to output " + mapOutputName);
                                }
                            }
                            else
                            {
                                // Turn labels off
                                _mapFunctions.SwitchLabels(mapOutputName, displayLabels, _logFile);
                            }
                        }
                        else
                        {
                            // User doesn't want to add the layer to the display.
                            _mapFunctions.RemoveLayer(mapOutputName, _logFile);
                        }
                    }

                    // Shouldn't need this as it's removed in the function ExportSelectionToCSV
                    //_mapFunctions.RemoveLayer(tempOutputLayerName, _logFile);
                    //if (_mapFunctions.FeatureclassExists(tempFCOutput))
                    //    _mapFunctions.DeleteFeatureclass(tempFCOutput, _logFile);

                    // Add to combined sites table as appropriate
                    // Function to take account of group by, order by and statistics fields.
                    if (mapCombinedSitesColumns != "" && combinedTableCreate)
                    {
                        FileFunctions.WriteLine(_logFile, "Extracting summary output for combined sites table ...");

                        intLineCount = _mapFunctions.ExportSelectionToCSV(tempMasterLayerName, strCombinedTable, mapCombinedSitesColumns, false, tempFCOutput, tempTableOutput, mapCombinedSitesGroupColumns,
                            mapCombinedSitesStatsColumns, mapCombinedSitesOrderColumns, includeArea, areaMeasureUnit, includeDistance, radiusText, targetLayer, _logFile, false, RenameColumns: true);

                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " row(s) added to summary output");

                        // Shouldn't need this as it's removed in the function ExportSelectionToCSV
                        //_mapFunctions.RemoveLayer(tempOutputLayerName, _logFile);
                        //if (_mapFunctions.FeatureclassExists(tempFCOutput))
                        //    _mapFunctions.DeleteFeatureclass(tempFCOutput, _logFile);
                    }

                    // Cleanup the temporary master layer.
//                    FileFunctions.WriteLine(_logFile, "Cleaning up temporary master layer");
                    _mapFunctions.RemoveLayer(tempMasterLayerName, _logFile);
                    _mapFunctions.DeleteFeatureclass(tempMasterOutput, _logFile);

                    // Clear the selection in the input layer.
 //                   FileFunctions.WriteLine(_logFile, "Clearing map selection");
                    _mapFunctions.ClearSelectedMapFeatures(mapLayerName, _logFile);
                    FileFunctions.WriteLine(_logFile, "Analysis complete");
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
