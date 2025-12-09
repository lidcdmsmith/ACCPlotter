// References:
// 

namespace ACCPlotter
{
    public class Command_About : System.Windows.Input.ICommand
    {
        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter)
        {
            return true;
        }

        public void Execute(object? parameter)
        {
            // Show About Message Box:
            System.Windows.MessageBox.Show
                (
                "ACC Plotter Tools for Civil 3D 2025\nVersion 1.0.0\nContact: David Li (lid@cdmsmith.com)",
                "Tool Information",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Information
                );
        }
    }
}
