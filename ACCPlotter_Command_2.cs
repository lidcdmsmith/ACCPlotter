// References:
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using Autodesk.Civil.DatabaseServices;
using Autodesk.Civil.ApplicationServices;
using System.IO;
using System.Windows.Media.Imaging;

namespace ACCPlotter
{
    // class    : command handler for the Ribbon Button 2
    public partial class Command_2 : System.Windows.Input.ICommand
    {
        public event EventHandler? CanExecuteChanged;

        // method   : CanExecute (called to check if command handler can execute)
        public bool CanExecute(object? parameter)
        {
            return true; // Always enabled
        }


        // method   : Execute (called when the ribbon button is clicked)
        public void Execute(object? parameter)
        {
            // place holder:
        }
    }
}
