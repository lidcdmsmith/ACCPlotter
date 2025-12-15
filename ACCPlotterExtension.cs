// References:
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System.IO;

using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

[assembly: ExtensionApplication(typeof(CDMSPlotting.PlottingCommands))]

namespace CDMSPlotting
{
    public class PlottingExtension : IExtensionApplication
    {
        public void Initialize()
        {
            // Initialization when Civil 3D starts
        }
        public void Terminate()
        {
            // Cleanup code when Civil 3D closes
        }
    }
}
