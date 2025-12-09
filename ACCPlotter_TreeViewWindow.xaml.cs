using System.IO;
using System.Windows;
using System.Collections.ObjectModel;
using System.Linq;



namespace ACCPlotter
{
    public partial class ACCPlotterFolderTreeViewWindow : Window
    {
        // Property: dynamic list of folder nodes for data binding to the TreeView
        public ObservableCollection<FolderNode> Folders { get; set; }

        public bool AnyFolderCheck()
        {
            foreach (var folder in Folders)
            {
                if (folder.IsChecked) 
                    return true;
            }
            return false;
        }

        // Folder check:
        // SubFolder check:

        


        // Constructor: Initializes the window and sets up the folder collection
        // projectFolderPath: The root folder path to build the tree from
        public ACCPlotterFolderTreeViewWindow(string projectFolderPath)
        {
            // Initialize UI from XAML file.
            InitializeComponent(); 

            Folders = new ObservableCollection<FolderNode>();

            // Populate TreeView
            //      Populate Folders
            //          Populate SubFolders
        }


        // Event Handler: Tree View Item Selection
        // Called when the selected item in the TreeView changes
        private void EventHandler1(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            // Handle selection change if needed:
            // If selected folder count is greater than 0, enable OK button
            // else disable it
        }


        // Event Handler: OK Button Click
        // Called when the OK button is clicked, closes the window
        private void EventHandler2(object sender, RoutedEventArgs e)
        {
            // Add selected folders to list
            // Then close the window
            Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {

        }
    }

    // FolderNode class for TreeView
    // Holds folder info and subfolders for the tree
    public class FolderNode
    {
        // Name of the folder (displayed in the TreeView)
        public string Name { get; set; } = "";

        // Full path of the folder
        public string FullPath { get; set; } = "";

        // Checkbox state for the selection in the UI
        public bool IsChecked { get; set; }
    }
}
