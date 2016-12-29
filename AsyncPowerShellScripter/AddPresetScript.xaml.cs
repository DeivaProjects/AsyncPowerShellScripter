using System.Windows;
using System.Windows.Controls;

namespace AsyncPowerShellScripter
{
    /// <summary>
    /// Interaction logic for AddPresetScript.xaml
    /// </summary>
    public partial class AddPresetScript : UserControl
    {
        private MainWindow _main;

        public AddPresetScript(MainWindow mainWindow)
        {
            InitializeComponent();
            _main = mainWindow;
            SaveButton.Click += SaveButton_Click;
            CancelButton.Click += CancelButton_Click;
            ScriptTitle.TextChanged += ScriptTitle_TextChanged;

            if (ScriptTitle.Text.Length == 0)
                SaveButton.IsEnabled = false;
        }

        private void ScriptTitle_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (ScriptTitle.Text.Length == 0)
                SaveButton.IsEnabled = false;
            else
                SaveButton.IsEnabled = true;
        }

        private void CancelButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Window parentWindow = Window.GetWindow((DependencyObject)sender);
            if (parentWindow != null)
            {
                parentWindow.Close();
            }
        }

        private void SaveButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            _main.SaveNewScript(ScriptTitle.Text);

            Window parentWindow = Window.GetWindow((DependencyObject)sender);
            if (parentWindow != null)
            {
                parentWindow.Close();
            }
        }
    }
}