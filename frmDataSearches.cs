
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











            // Go through each of the requested layers and carry out the relevant analysis. 
            List<string> strLayerNames = _toolConfig.MapLayers;
            List<string> strDisplayNames = _toolConfig.MapNames;
            List<string> strGISOutNames = _toolConfig.MapGISOutNames;
            List<string> strTableOutNames = _toolConfig.MapTableOutNames;
            List<string> strColumnList = _toolConfig.MapColumns;
            List<string> strGroupColumnList = _toolConfig.MapGroupByColumns;
            List<string> strStatsColumnList = _toolConfig.MapStatisticsColumns;
            List<string> strOrderColumnList = _toolConfig.MapOrderByColumns;
            List<string> strCriteriaList = _toolConfig.MapCriteria;
            List<bool> blIncludeAreas = _toolConfig.getMapIncludeAreas;
            List<bool> blIncludeDistances = _toolConfig.getMapIncludeDistances;
            List<bool> blIncludeRadii = _toolConfig.getMapIncludeRadii;
            List<string> strKeyColumns = _toolConfig.MapKeyColumns;
            List<string> strFormats = _toolConfig.MapFormats;
            List<bool> blKeepLayers = _toolConfig.MapKeepLayers;
            List<string> strOutputTypes = _toolConfig.MapOutputTypes;
            List<bool> blDisplayLabels = _toolConfig.MapDisplayLabels;
            List<string> strDisplayLayerFiles = _toolConfig.MapLayerFiles;
            List<bool> blOverwriteLabelDefaults = _toolConfig.MapOverwriteLabels;
            List<string> strLabelColumns = _toolConfig.MapLabelColumns;
            List<string> strLabelClauses = _toolConfig.MapLabelClauses;
            List<string> strMacroNames = _toolConfig.MapMacroNames;
            List<string> strCombinedSitesColumnList = _toolConfig.MapCombinedSitesColumns;
            List<string> strCombinedSitesGroupColumnList = _toolConfig.MapCombinedSitesGroupByColumns;
            List<string> strCombinedSitesStatsColumnList = _toolConfig.MapCombinedSitesStatsColumns;
            List<string> strCombinedSitesOrderColumnList = _toolConfig.MapCombinedSitesOrderByColumns;
            //List<string> strCombinedSitesCriteriaList = _toolConfig.MapCombinedSitesCriteria;

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




            foreach (string aLayer in SelectedLayers)
            {
                // Get all the settings relevant to this layer.
                int intIndex = strLayerNames.IndexOf(aLayer); // Finds the first occurrence. 
                string strDisplayName = strDisplayNames[intIndex];
                string strGISOutName = strGISOutNames[intIndex];
                string strTableOutName = strTableOutNames[intIndex];
                string strColumns = strColumnList[intIndex]; // Note there could be multiple columns.
                string strGroupColumns = strGroupColumnList[intIndex];
                string strStatsColumns = strStatsColumnList[intIndex];
                string strOrderColumns = strOrderColumnList[intIndex];
                string strCriteria = strCriteriaList[intIndex];
                bool blIncludeArea = blIncludeAreas[intIndex];
                bool blIncludeDistance = blIncludeDistances[intIndex];
                bool blIncludeRadius = blIncludeRadii[intIndex];

                string strKeyColumn = strKeyColumns[intIndex];
                string strFormat = strFormats[intIndex];
                bool blKeepLayer = blKeepLayers[intIndex];
                string strOutputType = strOutputTypes[intIndex];
                bool blDisplayLabel = blDisplayLabels[intIndex];
                string strDisplayLayer = strDisplayLayerFiles[intIndex];
                bool blOverwriteLabelDefault = blOverwriteLabelDefaults[intIndex];
                string strLabelColumn = strLabelColumns[intIndex];
                string strLabelClause = strLabelClauses[intIndex];
                string strMacroName = strMacroNames[intIndex];

                string strCombinedSitesColumns = strCombinedSitesColumnList[intIndex];
                string strCombinedSitesGroupColumns = strCombinedSitesGroupColumnList[intIndex];
                string strCombinedSitesStatsColumns = strCombinedSitesStatsColumnList[intIndex];
                string strCombinedSitesOrderColumns = strCombinedSitesOrderColumnList[intIndex];



                // Deal with wildcards in the output names.
                strGISOutName = StringFunctions.ReplaceSearchStrings(strGISOutName, reference, siteName, shortRef, subref, radius);
                strTableOutName = StringFunctions.ReplaceSearchStrings(strTableOutName, reference, siteName, shortRef, subref, radius);

                // Remove any illegal characters from the names.
                strGISOutName = StringFunctions.StripIllegals(strGISOutName, repChar);
                strTableOutName = StringFunctions.StripIllegals(strTableOutName, repChar);

                strStatsColumns = StringFunctions.AlignStatsColumns(strColumns, strStatsColumns, strGroupColumns);
                //if (blIncludeDistance && !strColumns.Contains("Distance") && !strGroupColumns.Contains("Distance"))
                //    strColumns = strColumns + ",Distance"; // Distance comes after grouping and hence should not be included in the stats columns.
                //if (blIncludeRadius && !strColumns.Contains("Radius") && !strGroupColumns.Contains("Radius"))
                //    strColumns = strColumns + ",Radius"; // as for Distance column, it comes after the grouping.
                strCombinedSitesStatsColumns = StringFunctions.AlignStatsColumns(strCombinedSitesColumns, strCombinedSitesStatsColumns, strCombinedSitesGroupColumns);

                // Create relevant output name. Note this is done whether or not the layer is eventually kept.
                string strShapeLayerName = strGISOutName;
                string strShapeOutputName = gisFolder + @"\" + strGISOutName; // output shapefile / feature class name. Note no extension to allow write to GDB
                string strTableOutputName = gisFolder + @"\" + strTableOutName + "." + strFormat.ToLower(); // output table name

                FileFunctions.WriteLine(_logFile, "Starting analysis for " + aLayer);

                // Do the selection. 
                FileFunctions.WriteLine(_logFile, "Selecting features using selected feature(s) from layer " + strLayerName + " ...");
                MapFunctions.SelectLayerByLocation(strDisplayName, strLayerName, aLogFile: _logFile);

                // Refine the selection if required.
                int featureCount = MapFunctions.CountSelectedLayerFeatures(strDisplayName, _logFile);
                if (featureCount > 0 && strCriteria != "")
                {
                    FileFunctions.WriteLine(_logFile, "Refining selection with criteria " + strCriteria + " ...");
                    blResult = MapFunctions.SelectLayerByAttributes(strDisplayName, strCriteria, "SUBSET_SELECTION", _logFile);
                    if (!blResult)
                    {
                        MessageBox.Show("Error selecting layer " + strDisplayName + " with criteria " + strCriteria + ". Please check syntax and column names (case sensitive)");
                        FileFunctions.WriteLine(_logFile, "Error refining selection on layer " + strDisplayName + " with criteria " + strCriteria + ". Please check syntax and column names (case sensitive)");
                        this.BringToFront();
                        this.Cursor = Cursors.Default;
                        MapFunctions.ToggleDrawing(true);
                        MapFunctions.ToggleTOC();
                        return;
                    }
                    //FileFunctions.WriteLine(_logFile, "Selection on " + strDisplayName + " refined");
                }

                // Write out the results - to shapefile first. Include distance if required.
                // Function takes account of output, group by and statistics fields.

                featureCount = MapFunctions.CountSelectedLayerFeatures(strDisplayName, _logFile);
                if (featureCount > 0)
                {
                    FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", featureCount) + " feature(s) found");
                    // Firstly take a copy of the full selection in a temporary file; This will be used to do the summaries on.

                    // Get the input FC type.
                    IFeatureClass pFC = MapFunctions.GetFeatureClassFromLayerName(strDisplayName, _logFile, false);
                    string strInFCType = MapFunctions.GetFeatureClassType(pFC);

                    // Get the buffer FC type.
                    pFC = MapFunctions.GetFeatureClassFromLayerName(strLayerName, _logFile, false);
                    string strBufferFCType = MapFunctions.GetFeatureClassType(pFC);

                    // If the input layer should be clipped to the buffer layer, do so now.
                    if (strOutputType == "CLIP")
                    {

                        if ((strInFCType == "polygon" & strBufferFCType == "polygon") ||
                            (strInFCType == "line" & (strBufferFCType == "line" || strBufferFCType == "polygon")))
                        {
                            // Clip
                            FileFunctions.WriteLine(_logFile, "Clipping selected features ...");
                            blResult = MapFunctions.ClipFeatures(strDisplayName, strOutputFile, strTempMasterOutput, aLogFile: _logFile); // Selected features in input, buffer FC, output.
                        }
                        else
                        {
                            // Copy
                            FileFunctions.WriteLine(_logFile, "Copying selected features ...");
                            blResult = MapFunctions.CopyFeatures(strDisplayName, strTempMasterOutput, aLogFile: _logFile);
                        }
                    }
                    // If the buffer layer should be clipped to the input layer, do so now.
                    else if (strOutputType == "OVERLAY")
                    {

                        if ((strBufferFCType == "polygon" & strInFCType == "polygon") ||
                            (strBufferFCType == "line" & (strInFCType == "line" || strInFCType == "polygon")))
                        {
                            // Clip
                            FileFunctions.WriteLine(_logFile, "Overlaying selected features ...");
                            blResult = MapFunctions.ClipFeatures(strOutputFile, strDisplayName, strTempMasterOutput, aLogFile: _logFile); // Selected features in input, buffer FC, output.
                        }
                        else
                        {
                            // Select from the buffer layer.
                            FileFunctions.WriteLine(_logFile, "Selecting features  ...");
                            MapFunctions.SelectLayerByLocation(strLayerName, strDisplayName, aLogFile: _logFile);

                            featureCount = MapFunctions.CountSelectedLayerFeatures(strLayerName, _logFile);
                            if (featureCount > 0)
                            {
                                FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", featureCount) + " feature(s) found");

                                // Copy the selection from the buffer layer.
                                FileFunctions.WriteLine(_logFile, "Copying selected features ... ");
                                blResult = MapFunctions.CopyFeatures(strLayerName, strTempMasterOutput, aLogFile: _logFile);
                            }
                            else
                            {
                                FileFunctions.WriteLine(_logFile, "No features selected");
                            }

                            // Clear the buffer layer selection.
                            MapFunctions.ClearSelectedMapFeatures(strLayerName, _logFile);
                        }
                    }
                    // If the input layer should be intersected with the buffer layer, do so now.
                    else if (strOutputType == "INTERSECT")
                    {

                        if ((strInFCType == "polygon" & strBufferFCType == "polygon") ||
                            (strInFCType == "line" & strBufferFCType == "line"))
                        {
                            // Intersect
                            FileFunctions.WriteLine(_logFile, "Intersecting selected features ...");
                            blResult = MapFunctions.IntersectFeatures(strDisplayName, strOutputFile, strTempMasterOutput, aLogFile: _logFile); // Selected features in input, buffer FC, output.
                        }
                        else
                        {
                            // Copy
                            FileFunctions.WriteLine(_logFile, "Copying selected features ...");
                            blResult = MapFunctions.CopyFeatures(strDisplayName, strTempMasterOutput, aLogFile: _logFile);
                        }
                    }
                    // Otherwise do a straight copy of the input layer.
                    else
                    {
                        // Copy
                        FileFunctions.WriteLine(_logFile, "Copying selected features ...");
                        blResult = MapFunctions.CopyFeatures(strDisplayName, strTempMasterOutput, aLogFile: _logFile);
                    }

                    if (!blResult)
                    {
                        MessageBox.Show("Cannot copy selection from " + strDisplayName + " to " + strTempMasterOutput + ". Please ensure this file is not open elsewhere");
                        FileFunctions.WriteLine(_logFile, "Cannot copy selection from " + strDisplayName + " to " + strTempMasterOutput + ". Please ensure this file is not open elsewhere");
                        this.BringToFront();
                        this.Cursor = Cursors.Default;
                        MapFunctions.ToggleDrawing(true);
                        MapFunctions.ToggleTOC();
                        return;
                    }

                    // Get the group name of the layer
                    string strGroupName = StringFunctions.GetGroupName(aLayer);
                    bool blNewLabelField = false;
                    
                    // If a label column is given
                    if (strLabelColumn != "")
                    {
                        // Check if the map label field exists. Create if necessary. 
                        if (!MapFunctions.FieldExists(strTempMasterOutput, strLabelColumn, aLogFile: _logFile) && addSelectedLayers.ToLower().Contains("with"))
                        {
                            // If not, create it and label.
                            MapFunctions.AddField(strTempMasterOutput, strLabelColumn, esriFieldType.esriFieldTypeInteger, 10, _logFile);
                            blNewLabelField = true;
                        }

                        // Add labels as required
                        if (blNewLabelField || (overwriteLabels.ToLower() != "no" && blOverwriteLabelDefault && addSelectedLayers.ToLower().Contains("with") && strLabelColumn != ""))
                        // Either we  have a new label field, or we want to overwrite the labels and are allowed to.
                        {
                            // Add relevant labels. 
                            if (overwriteLabels.ToLower().Contains("layer")) // Reset each layer to 1.
                            {
                                FileFunctions.WriteLine(_logFile, "Resetting label counter ...");
                                startLabel = 1;
                                FileFunctions.WriteLine(_logFile, "Adding map labels ...");
                                MapFunctions.AddIncrementalNumbers(strTempMasterOutput, strLabelColumn, strKeyColumn, startLabel, _logFile);
                            }

                            else if (resetGroups && strGroupName != "")
                            {
                                // Increment within but reset between groups. Note all group labels are already initialised as 1.
                                // Only triggered if a group name has been found.

                                int intGroupIndex = liGroupNames.IndexOf(strGroupName);
                                int intGroupLabel = liGroupLabels[intGroupIndex];
                                FileFunctions.WriteLine(_logFile, "Adding map labels ...");
                                intGroupLabel = MapFunctions.AddIncrementalNumbers(strTempMasterOutput, strLabelColumn, strKeyColumn, intGroupLabel, _logFile);
                                intGroupLabel++;
                                liGroupLabels[intGroupIndex] = intGroupLabel; // Store the new max label.
                            }
                            else
                            {
                                // There is no group or groups are ignored, or we are not resetting. Use the existing max label number.
                                startLabel = maxLabel;

                                FileFunctions.WriteLine(_logFile, "Adding map labels ...");
                                maxLabel = MapFunctions.AddIncrementalNumbers(strTempMasterOutput, strLabelColumn, strKeyColumn, startLabel, _logFile);
                                maxLabel++; // the new start label for incremental labeling
                            }
                        }
                    }

                    //string strTempShapeFile = tempFolder + @"\TempShapes_" + strUserID + ".shp";
                    //string strTempShapeOutput = tempFolder + @"\" + strTempOutput + ".shp";
                    //string strTempDBFOutput = tempFolder + @"\" + strTempOutput + "DBF.dbf";
                    string radiusText = "none";
                    if (blIncludeRadius) radiusText = radius; // Only include radius if requested.

                    // Only export if the user has specified columns.
                    int intLineCount = -999;
                    if (strColumns != "")
                    {
                        // Write out the results to table as appropriate.
                        bool blIncHeaders = false;
                        if (strFormat.ToLower() == "csv") blIncHeaders = true;

                        FileFunctions.WriteLine(_logFile, "Extracting summary information ...");

                        intLineCount = MapFunctions.ExportSelectionToCSV(tempMasterLayerName, strTableOutputName, strColumns, blIncHeaders, tempFCOutput, tempTableOutput,
                            strGroupColumns, strStatsColumns, strOrderColumns, blIncludeArea, areaMeasureUnit, blIncludeDistance, radiusText, targetLayer, _logFile, RenameColumns: true);

                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " record(s) exported");
                    }

                    // Copy to permanent layer as appropriate
                    if (blKeepLayer)
                    {
                        // Keep the layer - write to permanent file (note this is not the summarised layer).
                        FileFunctions.WriteLine(_logFile, "Copying selected GIS features to " + strShapeLayerName + ".shp ...");
                        MapFunctions.CopyFeatures(tempMasterLayerName, strShapeOutputName, aLogFile: _logFile);

                        // If the layer is to be added to the map
                        if (addSelectedLayers.ToLower() != "no")
                        {
                            // Add the permanent layer to the map (by moving it to the group layer)
                            MapFunctions.MoveToGroupLayer(groupLayerName, MapFunctions.GetLayer(strShapeLayerName, _logFile), _logFile);

                            // If there is a layer file
                            if (strDisplayLayer != "")
                            {
                                string strDisplayLayerFile = _layerFolder + @"\" + strDisplayLayer;
                                MapFunctions.ChangeLegend(strShapeLayerName, strDisplayLayerFile, blDisplayLabel, _logFile);
                            }

                            FileFunctions.WriteLine(_logFile, "Output " + strShapeLayerName + " added to display");

                            // If labels are to be displayed
                            if (addSelectedLayers.ToLower().Contains("with ") && blDisplayLabel)
                            {
                                // Translate the label string.
                                if (strLabelClause != "" && strDisplayLayer == "") // Only if we don't have a layer file.
                                {
                                    List<string> strLabelOptions = strLabelClause.Split('$').ToList();
                                    string strFont = strLabelOptions[0].Split(':')[1];
                                    double dblSize = double.Parse(strLabelOptions[1].Split(':')[1]); // Needs error trapping
                                    int intRed = int.Parse(strLabelOptions[2].Split(':')[1]); // Needs error trapping
                                    int intGreen = int.Parse(strLabelOptions[3].Split(':')[1]);
                                    int intBlue = int.Parse(strLabelOptions[4].Split(':')[1]);
                                    string strOverlap = strLabelOptions[5].Split(':')[1];
                                    MapFunctions.AnnotateLayer(strShapeLayerName, "[" + strLabelColumn + "]", strFont, dblSize,
                                        intRed, intGreen, intBlue, strOverlap, aLogFile: _logFile);
                                    FileFunctions.WriteLine(_logFile, "Labels added to output " + strShapeLayerName);
                                }
                                else if (strLabelColumn != "" && strDisplayLayer == "")
                                {
                                    MapFunctions.AnnotateLayer(strShapeLayerName, "[" + strLabelColumn + "]", aLogFile: _logFile);
                                    FileFunctions.WriteLine(_logFile, "Labels added to output " + strShapeLayerName);
                                }
                            }
                            else
                            {
                                // Turn labels off
                                MapFunctions.SwitchLabels(strShapeLayerName, blDisplayLabel, _logFile);
                            }
                        }
                        else
                        {
                            // User doesn't want to add the layer to the display.
                            MapFunctions.RemoveLayer(strShapeLayerName, _logFile);
                        }
                    }

                    // Shouldn't need this as it's removed in the function ExportSelectionToCSV
                    //MapFunctions.RemoveLayer(tempOutputLayerName, _logFile);
                    //if (MapFunctions.FeatureclassExists(tempFCOutput))
                    //    MapFunctions.DeleteFeatureclass(tempFCOutput, _logFile);

                    // Add to combined sites table as appropriate
                    // Function to take account of group by, order by and statistics fields.
                    if (strCombinedSitesColumns != "" && combinedTableCreate)
                    {
                        FileFunctions.WriteLine(_logFile, "Extracting summary output for combined sites table ...");

                        intLineCount = MapFunctions.ExportSelectionToCSV(tempMasterLayerName, strCombinedTable, strCombinedSitesColumns, false, tempFCOutput, tempTableOutput, strCombinedSitesGroupColumns,
                            strCombinedSitesStatsColumns, strCombinedSitesOrderColumns, blIncludeArea, areaMeasureUnit, blIncludeDistance, radiusText, targetLayer, _logFile, false, RenameColumns: true);

                        FileFunctions.WriteLine(_logFile, string.Format("{0:n0}", intLineCount) + " row(s) added to summary output");

                        // Shouldn't need this as it's removed in the function ExportSelectionToCSV
                        //MapFunctions.RemoveLayer(tempOutputLayerName, _logFile);
                        //if (MapFunctions.FeatureclassExists(tempFCOutput))
                        //    MapFunctions.DeleteFeatureclass(tempFCOutput, _logFile);
                    }

                    // Cleanup the temporary master layer.
//                    FileFunctions.WriteLine(_logFile, "Cleaning up temporary master layer");
                    MapFunctions.RemoveLayer(tempMasterLayerName, _logFile);
                    MapFunctions.DeleteFeatureclass(tempMasterOutput, _logFile);

                    // Clear the selection in the input layer.
 //                   FileFunctions.WriteLine(_logFile, "Clearing map selection");
                    MapFunctions.ClearSelectedMapFeatures(strDisplayName, _logFile);
                    FileFunctions.WriteLine(_logFile, "Analysis complete");
                }
                else
                {
                    FileFunctions.WriteLine(_logFile, "No features found");
                }

                // Trigger the macro if one exists
                if (strMacroName != "") // Only if we have a macro to trigger.
                {
                    FileFunctions.WriteLine(_logFile, "Executing vbscript macro ...");

                    Process scriptProc = new Process();
                    //scriptProc.StartInfo.FileName = @"C:\Windows\SysWOW64\cscript.exe";
                    scriptProc.StartInfo.FileName = @"cscript.exe";
                    scriptProc.StartInfo.WorkingDirectory = FileFunctions.GetDirectoryName(strMacroName); //<---very important
                    scriptProc.StartInfo.UseShellExecute = true;
                    scriptProc.StartInfo.Arguments = string.Format(@"//B //Nologo {0} {1} {2} {3}", "\"" + strMacroName + "\"", "\"" + gisFolder + "\"", "\"" + strTableOutName + "." + strFormat.ToLower() + "\"", "\"" + strTableOutName + ".xlsx" + "\"");
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
            MapFunctions.UpdateTOC();
            MapFunctions.ToggleTOC();
            MapFunctions.ToggleDrawing(true);
            MapFunctions.SetContentsView();
            if (bufferSize != "0" && blKeepBuffer)
                MapFunctions.ZoomToLayer(strLayerName, _logFile);



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

                    string strColumns = siteColumn;
                    if (!string.IsNullOrEmpty(orgColumn))
                    {
                        if (!string.IsNullOrEmpty(strColumns))
                            strColumns = strColumns + "," + orgColumn;
                        else
                            strColumns = orgColumn;
                    }

                    string searchQuery = "SELECT " + strColumns + " from " + strTable + " WHERE LCASE(" + _toolConfig.RefColumn + ") = " + '"' + txtSearch.Text.ToLower() + '"';
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
