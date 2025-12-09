// References:
using Autodesk.Aec.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Civil.ApplicationServices;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Windows;
using System.Diagnostics;
using System.IO;
using System.Printing;
using System.Security.Cryptography;
using System.Windows.Annotations;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Media.Media3D;
using static System.Net.WebRequestMethods;

namespace ACCPlotter
{
    // class    : command handler for the Ribbon Button 1
    public partial class Command_1 : System.Windows.Input.ICommand
    {
        // property: CanExecuteChanged (event required by ICommand interface)
        public event EventHandler? CanExecuteChanged;

        // method   : CanExecute (called to check if command handler can execute)
        public bool CanExecute(object? parameter)
        {
            return true; // Always enabled
        }

        // method   : Execute (called when the ribbon button is clicked)
        public void Execute(object? parameter)
        {
            // Editor
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            // get project folder, scan list, folder list, cad file list, cad template:
            string? projectFolderPath = GetProjectFolderPath();

            //string? dwgFolderScanListPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "dwgFolderScanList.txt");
            string? cadTemplateFilePath = GetDwgCadFileTemplatePath();

            //if (projectFolderPath is null || dwgFolderScanListPath is null || cadTemplateFilePath is null) return;
            if (projectFolderPath is null || cadTemplateFilePath is null) return;

            //List<string> dwgFolderScanList = GetDwgFolderScanList(dwgFolderScanListPath);
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

            // Prepare log colection
            var logRows = new List<List<string>>();

            // Iterate through each project CAD file and perform plotting operations
            foreach (string cadFilePath in cadFilePathList)
            {
                using (Database database = new Database(false, true))
                {
                    DatabaseReadDwgFile(database, cadFilePath);

                    // start transactions
                    using (Transaction transaction = database.TransactionManager.StartTransaction())
                    {
                        //  get layout dictionary
                        DBDictionary layoutDict = (DBDictionary)(transaction.GetObject(database.LayoutDictionaryId, OpenMode.ForWrite));

                        // create plot setting validators for applying plot settings
                        PlotSettingsValidator psv = PlotSettingsValidator.Current;

                        // get plot settings from template file
                        PlotSettings? templatePlotSettings = PlotSettingsExtract(cadTemplateFilePath);
                        if (templatePlotSettings is null) return;

                        // iterate through layouts and apply plot settings from template
                        foreach (DBDictionaryEntry layoutDictEntry in layoutDict)
                        {
                            Layout layout = (Layout)(transaction.GetObject(layoutDictEntry.Value, OpenMode.ForWrite));

                            if (layout.ModelType || layout is null) 
                                continue;

                            ed.WriteMessage($"\nApplying plot settings to {Path.GetFileName(cadFilePath)} - {layout.LayoutName}");

                            PlotSettingsApply(layout, templatePlotSettings);

                            // Log applied settings:
                            logRows.Add(GetPlotSettingsLog(cadFilePath, layout));
                        }

                        // commit transaction
                        transaction.Commit();
                    }
                    // save dataase after committing transaction
                    database.SaveAs(cadFilePath, DwgVersion.Current);
                    // database.Dispose(); automatically called by "using" statement
                }
            }

            // Write log to CSV:
            string tempFolder = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            tempFolder = Path.Combine(tempFolder, "AppData", "Local", "Temp");

            string projectFolderName = !string.IsNullOrEmpty(projectFolderPath)
                ? new DirectoryInfo(projectFolderPath).Name
                : "UnknownProject";

            string dateString = DateTime.Now.ToString("yyyy-MM-dd_H-mm-ss");

            string logFileName = $"ACCPlotter_Log_{projectFolderName}_{dateString}.csv";
            string logPath = Path.Combine(tempFolder, logFileName);

            using (var writer = new StreamWriter(logPath))
            {
                // Write header
                writer.WriteLine("dwgFile, LayoutName, PlotConfigurationName, PlotSettingsName, PlotPaperSize, PlotOrigin, PlotCentered," +
                    "PlotPaperMargins, PlotPaperUnits, ScaleLineweights, ShowPlotStyles, AnnoAllVisible, Annotative, CurrentStyleSheet, DrawViewportsFirst," +
                    "PlotAsRaster, PlotHidden, PlotType, PlotViewName, PlotViewportBorders, PlotWindowArea, PlotWireframe, ShadePlot, ShadePlotCustomDpi," +
                    "ShadePlotResLevel, PrintLineweights, PlotTransparency, PlotPlotStyles, TabOrder, UseStandardScale, PlotRotation, PaperOrientation");

                foreach (var row in logRows)
                writer.WriteLine(string.Join(",", row.Select(v => $"\"{v}\"")));
            }
        }

        // helper   : DatabaseReadDwgFile
        // inputs   : database, cad file path
        // returns  : void (maybe return bool for success/failure?)
        public static void DatabaseReadDwgFile(Database database, string cadFilePath)
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

        // helper   : GetProjectFolderPath
        // inputs   : none
        // returns  : selected project folder path or null
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

        // helper   : GetDwgFolderScanList
        // inputs   : scan list file path
        // returns  : list of folder names to scan for dwg files
        public static List<string> GetDwgFolderScanList(string scanListFilePath)
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

        // helper   : GetDwgCadFilePathList
        // inputs   : project folder path, dwg folder scan list:
        // returns  : list of cad file paths
        public static List<string> GetDwgCadFilePathList(string projectFolderPath, List<string> dwgFolderScanList)
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

        // helper   : GetDwgCadFileTemplatePath
        // inputs   : optional filter string
        // returns  : selected cad file path or null
        public static string? GetDwgCadFileTemplatePath(string filter = "DWG and DWT Files (*.dwg;*.dwt)|*.dwg;*.dwt")
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

        // helper   : ExtractPlotSettings
        // inputs   : Layout
        // returns  : PlotSettings cloned from layout (pc3, media/paper size, ctb/stb, plot area, scale, offset & centering, rotation, qualtiy, to file/device)
        // comments : assumes first non-model layout found is the desired one
        public static PlotSettings? PlotSettingsExtract(string cadTemplateFilePath)
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

        // helper   : ApplyPlotSettings
        // inputs   : target Layout, source PlotSettings
        // returns  : void (maybe return bool for success/failure?)
        public static void PlotSettingsApply(Layout targetLayout, PlotSettings sourceSettings)
        {
            PlotSettingsValidator psv = PlotSettingsValidator.Current;
            psv.RefreshLists(targetLayout);

            // Copy relevant properties
            targetLayout.CopyFrom(sourceSettings);
        }

        // helper   : GetPlotSettingsLog
        // inputs   : layout, list
        // returns  : list
        public static List<string> GetPlotSettingsLog(string cadFilePath, Layout layout)
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
    }
}
