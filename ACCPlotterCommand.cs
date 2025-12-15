// References:
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System.IO;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: CommandClass(typeof(CDMSPlotting.PlottingCommands))]

namespace CDMSPlotting
{
    public class PlottingCommands()
    {
        [CommandMethod("CDMSPLOTTEMPLATE")]
        public void PlotPdfsTemplate()
        {
            #region InitialSetup
            Editor ed = AcadApp.DocumentManager.MdiActiveDocument.Editor;

            string? projectFolderPath = GetProjectFolderPath();

            string cadFileFilter = "DWG and DWT Files (*.dwg;*.dwt)|*.dwg;*.dwt";

            string? cadTemplateFilePath = GetDwgCadFileTemplatePath(cadFileFilter);

            if (projectFolderPath is null || cadTemplateFilePath is null) return;

            List<string> dwgFolderScanList = new List<string>
            {
                "01 General (G)",
                "02 Civil (C)",
                "03 Architectural (A)",
                "04 Structural (S)",
                "05 Process Mechanical",
                "06 HVAC (H)",
                "07 Plumbing (P)",
                "08 Fire Protection (F)",
                "09 Electrical (E)",
                "10 Automation (I)",
                "11 Plant3D",
                "12 Survey",
                "01 Gen",
                "02 Civil",
                "03 Arch",
                "04 Struc",
                "05 PMech",
                "06 HVAC",
                "07 PLBG",
                "08 FP",
                "09 Elec",
                "10 Autom",
                "11 P3D",
                "12 Survey"
            };

            List<string> cadFilePathList = GetDwgCadFilePathList(projectFolderPath, dwgFolderScanList);

            var logRows = new List<List<string>>();
            #endregion

            #region DWG Files Processing Loop
            foreach (string cadFilePath in cadFilePathList)
            {
                using (Database database = new Database(false, true))
                {
                    DatabaseReadDwgFile(database, cadFilePath);
            #endregion

                    #region Start Transaction
                    using (Transaction transaction = database.TransactionManager.StartTransaction())
                    {
                    #endregion

                        #region Layout and PlotSettings Dictionaries Access
                        DBDictionary layoutDict = (DBDictionary)(transaction.GetObject(database.LayoutDictionaryId, OpenMode.ForWrite));

                        PlotSettingsValidator psv = PlotSettingsValidator.Current;

                        // PlotSettings? templatePlotSettings = PlotSettingsExtract(cadTemplateFilePath);

                        if (templatePlotSettings is null) return;

                        DBDictionary plotSettingsDict = (DBDictionary)(transaction.GetObject(database.PlotSettingsDictionaryId, OpenMode.ForWrite));
                        #endregion

                        #region DWG Layouts Processing Loop
                        foreach (DBDictionaryEntry layoutDictEntry in layoutDict)
                        {
                        #endregion

                            #region Apply Plot Settings from Template
                            Layout layout = (Layout)(transaction.GetObject(layoutDictEntry.Value, OpenMode.ForWrite));

                            if (layout.ModelType || layout is null) continue;

                            ed.WriteMessage($"\nApplying plot settings to {Path.GetFileName(cadFilePath)} - {layout.LayoutName}");

                            string pageSetupName = $"PlotTemplate_{layout.LayoutName}";
                            
                            if (plotSettingsDict.Contains(templatePlotSettings.PlotSettingsName))
                            {
                                ObjectId existingPageSetupId = plotSettingsDict.GetAt(pageSetupName);

                                using (PlotSettings existingPageSetup = (PlotSettings)transaction.GetObject(existingPageSetupId, OpenMode.ForWrite))
                                    existingPageSetup.Erase();
                            }

                            using (PlotSettings newPageSetup = new PlotSettings(false))
                            {
                                newPageSetup.CopyFrom(templatePlotSettings);

                                string plotDeviceName = templatePlotSettings.PlotConfigurationName; // accesses pc3 file name

                                string mediaName = templatePlotSettings.CanonicalMediaName;

                                psv.RefreshLists(newPageSetup);

                                psv.SetPlotConfigurationName(newPageSetup, plotDeviceName, mediaName);

                                newPageSetup.AddToPlotSettingsDictionary(database);

                                layout.PlotSettingsName = pageSetupName;
                            }
                            logRows.Add(GetPlotSettingsLog(cadFilePath, layout));
                        }
                        #endregion

                        #region Commit Transaction
                        transaction.Commit();
                    }
                    database.SaveAs(cadFilePath, DwgVersion.Current);
                }
                #endregion

                #region Write Log to CSV
                WriteLogToCsv(logRows);
            }
            #endregion
        }
        #region PlotPdfsTemplate Private Methods

        /// <summary>
        /// Reads a DWG file into the provided Database object with error handling.
        /// </summary>
        private static void DatabaseReadDwgFile(Database database, string cadFilePath)
        {
            try
            {
                database.ReadDwgFile(cadFilePath, FileOpenMode.OpenForReadAndAllShare, true, "");
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                System.Windows.MessageBox.Show(
                    $"AutoCAD error reading DWG file:\n{cadFilePath}\n{ex.Message}",
                    "DWG File Read Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show(
                $"System error reading DWG file:\n{cadFilePath}\n{ex.Message}",
                "DWG File Read Error",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Gets the project folder path from a folder browser dialog.
        /// </summary>
        public static string? GetProjectFolderPath()
        {
            string? projectFolderPath = null;

            using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                dialog.Description = "Select Project Folder";
                dialog.UseDescriptionForTitle = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                    projectFolderPath = dialog.SelectedPath;

                if (string.IsNullOrEmpty(projectFolderPath))
                    projectFolderPath = null;
            }
            return projectFolderPath;
        }

        /// <summary>
        /// (NOT USED) Gets a list of folder names to scan for DWG files from a specified scan list file.
        /// FolderScanList currently hard coded in PlotPdfsTemplate().
        /// </summary>
        private static List<string> GetDwgFolderScanList(string scanListFilePath)
        {
            List<string> scanList = new List<string>();

            if (System.IO.File.Exists(scanListFilePath))
            {
                // Read all lines from the scan list file
                string[] scanListLines = System.IO.File.ReadAllLines(scanListFilePath);
                scanList.Clear();
                foreach (var line in scanListLines)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                        scanList.Add(line);
                }
            }
            else
            {
                // message to inform empty scan list or missing file
                System.Windows.MessageBox.Show(
                $"System error reading sscan list file:\n{scanListFilePath}.",
                "Scan List File Read Error",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
            }
            return scanList;
        }

        /// <summary>
        ///  Gets a list of DWG CAD file paths from specified project folder and scan list.
        /// </summary>
        private static List<string> GetDwgCadFilePathList(string projectFolderPath, List<string> dwgFolderScanList)
        {
            // Initialize list to hold CAD file paths
            List<string> cadFilePaths = new List<string>();

            // Scan dwg files and add to path list:
            DirectoryInfo projectDir = new DirectoryInfo(projectFolderPath);

            foreach (DirectoryInfo subFolder in projectDir.GetDirectories())
            {
                if (dwgFolderScanList.Contains(subFolder.Name))
                {
                    DirectoryInfo caddFolder = new DirectoryInfo(Path.Combine(subFolder.FullName, "10 CADD"));

                    if (caddFolder.Exists)
                    {
                        foreach (FileInfo dwgFile in caddFolder.GetFiles("*.dwg"))
                            cadFilePaths.Add(dwgFile.FullName);
                    }
                }
            }
            return cadFilePaths;
        }

        /// <summary>
        /// Gets the DWG CAD template file path from an open file dialog.
        /// </summary>
        private static string? GetDwgCadFileTemplatePath(string filter)
        {
            using (var dialog = new System.Windows.Forms.OpenFileDialog())
            {
                dialog.Title = "Select CAD Template File";
                dialog.Filter = filter;
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                if (dialog.ShowDialog() == DialogResult.OK)
                    return dialog.FileName;
            }
            return null;
        }

        /// <summary>
        /// Extracts the <see cref="PlotSettings"/> from the first non-model layout in the specified CAD template file.
        /// </summary>
        private static PlotSettings? PlotSettingExtract(string cadTemplateFilePath)
        {
            // create new cad database:
            using (Database templateDatabase = new Database(false, true))
            {
                DatabaseReadDwgFile(templateDatabase, cadTemplateFilePath);
                using (Transaction transaction = templateDatabase.TransactionManager.StartTransaction())
                {
                    DBDictionary layoutDict = (DBDictionary)(transaction.GetObject(templateDatabase.LayoutDictionaryId, OpenMode.ForRead));
                    foreach (DBDictionaryEntry layoutDictEntry in layoutDict)
                    {
                        Layout layout = (Layout)(transaction.GetObject(layoutDictEntry.Value, OpenMode.ForRead));
                        if (layout.ModelType || layout is null)
                            continue;

                        PlotSettings plotSettings = layout; // Layout inherits from PlotSettings
                        return (PlotSettings)layout.Clone();
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Extracts the <see cref="PlotSettings"/> from the first non-model layout in the specified CAD template file.
        /// </summary>
        private static PlotSettings? PplotSettingsExtract(string cadTemplateFilePath)
        {
            // create new cad database:
            using (Database templateDatabase = new Database(false, true))
            {
                DatabaseReadDwgFile(templateDatabase, cadTemplateFilePath);
                using (Transaction transaction = templateDatabase.TransactionManager.StartTransaction())
                {
                    DBDictionary layoutDict = (DBDictionary)(transaction.GetObject(templateDatabase.LayoutDictionaryId, OpenMode.ForRead));
                    foreach (DBDictionaryEntry layoutDictEntry in layoutDict)
                    {
                        Layout layout = (Layout)(transaction.GetObject(layoutDictEntry.Value, OpenMode.ForRead));
                        if (layout.ModelType || layout is null)
                            continue;

                        PlotSettings plotSettings = layout; // Layout inherits from PlotSettings
                        return (PlotSettings)layout.Clone();
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// (NOT USED) Applies the extracted plot settings to the target layout.
        /// </summary>
        private static void PlotSettingsApply(Layout targetLayout, PlotSettings sourceSettings)
        {
            PlotSettingsValidator psv = PlotSettingsValidator.Current;
            psv.RefreshLists(targetLayout);

            // Copy relevant properties
            targetLayout.CopyFrom(sourceSettings);
        }

        /// <summary>
        /// Gets a log row of plot settings from the specified layout.
        /// </summary>
        /// <returns></returns>
        private static List<string> GetPlotSettingsLog(string cadFilePath, Layout layout)
        {
            var row = new List<string>
        {
            Path.GetFileName(cadFilePath),
            layout.LayoutName,
            layout.PlotConfigurationName,
            layout.PlotSettingsName,

            layout.PlotPaperSize.ToString(),

            layout.PlotOrigin.ToString(),
            layout.PlotCentered.ToString(),
            layout.PlotPaperMargins.ToString(),

            layout.PlotPaperUnits.ToString(),
            layout.ScaleLineweights.ToString(),

            layout.ShowPlotStyles.ToString(),

            layout.AnnoAllVisible.ToString(),
            layout.Annotative.ToString(),

            layout.CurrentStyleSheet,

            layout.DrawViewportsFirst.ToString(),

            layout.PlotAsRaster.ToString(),
            layout.PlotHidden.ToString(),
            layout.PlotType.ToString(),

            layout.PlotViewName,
            layout.PlotViewportBorders.ToString(),
            layout.PlotWindowArea.ToString(),
            layout.PlotWireframe.ToString(),

            layout.ShadePlot.ToString(),
            layout.ShadePlotCustomDpi.ToString(),
            layout.ShadePlotResLevel.ToString(),

            layout.PrintLineweights.ToString(),
            layout.PlotTransparency.ToString(),
            layout.PlotPlotStyles.ToString(),

            layout.TabOrder.ToString(),
            layout.UseStandardScale.ToString(),

            layout.PlotRotation.ToString(),
            layout.PaperOrientation.ToString(),
        };
            return row;
        }

        /// <summary>
        /// Writes the provided log data to a CSV file in the user's temporary folder.
        /// </summary>
        private static void WriteLogToCsv(List<List<string>> logRows)
        {
            string tempFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            tempFolder = Path.Combine(tempFolder, "AppData", "Local", "Temp");

            string dateString = DateTime.Now.ToString("yyyy-MM-dd_H-mm-ss");

            string logFileName = $"ACCPlotter_Log_{dateString}.csv";
            string logPath = Path.Combine(tempFolder, logFileName);

            using (var writer = new StreamWriter(logPath))
            {
                writer.WriteLine("dwgFile, LayoutName, PlotConfigurationName, PlotSettingsName, PlotPaperSize, PlotOrigin, PlotCentered," +
                    "PlotPaperMargins, PlotPaperUnits, ScaleLineweights, ShowPlotStyles, AnnoAllVisible, Annotative, CurrentStyleSheet, DrawViewportsFirst," +
                    "PlotAsRaster, PlotHidden, PlotType, PlotViewName, PlotViewportBorders, PlotWindowArea, PlotWireframe, ShadePlot, ShadePlotCustomDpi," +
                    "ShadePlotResLevel, PrintLineweights, PlotTransparency, PlotPlotStyles, TabOrder, UseStandardScale, PlotRotation, PaperOrientation");

                foreach (var row in logRows)
                    writer.WriteLine(string.Join(",", row.Select(v => $"\"{v}\"")));
            }
        }

        #endregion
    }
}
