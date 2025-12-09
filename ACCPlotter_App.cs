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
using System.Security.Cryptography.X509Certificates;

namespace ACCPlotter
{
    public class App : IExtensionApplication
    {
        public void Initialize()                                                        // Called when plugin is loaded (subscribe to ribbon initialization event).
        {
            // SubscribeToRibbonInitialization();
        }
        public void Terminate()                                                         // Called when the plugin is unloaded (unsubscribe from ribbon initialization event).
        {
            // UnsubscribeFromRibbonInitialization();
        }

        // Ribbon Tool:
        // public void SubscribeToRibbonInitialization()                                   // Subscribe "+=" to the ribbon initialization event: 
        // {
        //     ComponentManager.ItemInitialized += ComponentManager_RibbonInitialized;
        // }
        // public void UnsubscribeFromRibbonInitialization()                               // Unsubscribe "-=" from the ribbon initialization event;
        // {
        //     ComponentManager.ItemInitialized -= ComponentManager_RibbonInitialized;
        // }
        // private void ComponentManager_RibbonInitialized(object? sender, EventArgs e)    // Handles the ribbon initialization event
        // {
        //     ComponentManager.ItemInitialized -= ComponentManager_RibbonInitialized;     // Unsubscribe after the ribbon is initialized to ensure this only runes once:
        //     AddRibbonTab();
        // }
        // public void AddRibbonTab()                                                      // Method to create custom ribbon tab, panel, and buttons
        // {
        //     string thisAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
        // 
        //     System.Windows.Application.ResourceAssembly = typeof(App).Assembly;
        // 
        //     RibbonControl ribbonControl = ComponentManager.Ribbon;                      // Get the ribbon control:
        //     RibbonTab ribbonTab = new RibbonTab                                         // Create custom tab:
        //     {
        //         Title = "ACC Tools",
        //         Id = "{03620D0D-867A-45A3-9721-712335D8AB56}"                           // Unique identifier generated for the tab
        //     };
        //     ribbonControl.Tabs.Add(ribbonTab);                                          // Add tab to ribbon control
        // 
        //     RibbonPanelSource ribbonPanelSource = new RibbonPanelSource                 // Create main panel source:
        //     {
        //         Title = "Plotting"
        //     };
        // 
        //     RibbonPanel ribbonPanel = new RibbonPanel                                   // Create main panel:
        //     {
        //         Source = ribbonPanelSource
        //     };
        //     ribbonTab.Panels.Add(ribbonPanel);                                          // Add panel to tab:
        // 
        //     // Create a button "ribbonButton1" and assign the command handler
        //     Uri iconUri1 = new Uri(                                                     // Load icon image from resources
        //         "pack://application:,,,/Resources/icon1.ico",
        //         UriKind.Absolute);
        //     BitmapImage iconImage1 = new BitmapImage(iconUri1);
        // 
        //     RibbonButton ribbonButton1 = new RibbonButton
        //     {
        //         Text = "ACC Plotter 1",
        //         ShowText = true,
        //         ShowImage = true,
        //         Orientation = System.Windows.Controls.Orientation.Vertical,
        //         LargeImage = iconImage1,
        //         Size = RibbonItemSize.Large,
        //     };
        //     ribbonButton1.CommandHandler = new Command_1();                             // Assign the command handler
        // 
        // 
        //     // Create a button "ribbonButton2" and assign the command handler
        //     Uri iconUriAbout = new Uri(
        //     "pack://application:,,,/Resources/icon_about.ico",                      // Load icon image from resources
        //     UriKind.Absolute);
        //     BitmapImage iconImageAbout = new BitmapImage(iconUriAbout);
        // 
        //     RibbonButton ribbonButtonAbout = new RibbonButton
        //     {
        //         Text = "About",
        //         ShowText = true,
        //         ShowImage = true,
        //         Orientation = System.Windows.Controls.Orientation.Vertical,
        //         LargeImage = iconImageAbout,
        //         Size = RibbonItemSize.Standard,
        //     };
        //     ribbonButtonAbout.CommandHandler = new Command_About();                     // Assign the command handler
        // 
        //     // Setup main, sub panels layout:
        //     // create row panel, add button to it:
        //     RibbonRowPanel subPanelRow = new RibbonRowPanel();
        //     subPanelRow.Items.Add(ribbonButtonAbout);
        // 
        //     // add main panel items:
        //     ribbonPanelSource.Items.Add(ribbonButton1);
        // 
        //     //add panel break to start slide-out section:
        //     ribbonPanelSource.Items.Add(new RibbonPanelBreak());
        // 
        //     // Add slide out panel row (with about button) to main panel:
        //     ribbonPanelSource.Items.Add(subPanelRow);
        // }


        // Command Tool:
        [CommandMethod("ACCPLOTTER")]
        public void ACCPlotter()                                                   // Command to force initialization of the plugin
        {
            // Create an instance and run main logic:
            var cmd = new Command_1();
            cmd.Execute(null);
        }
    }
}
