using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;

namespace AsyncPowerShellScripter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Runspace runSpace;
        private PipelineExecutor pipelineExecutor;
        private BackgroundWorker worker;
        private List<PresetScript> presetScriptItems;
        private bool savingNewScript;
        private string savingNewScriptName;

        public MainWindow()
        {
            InitializeComponent();

            // Set MainWindow icon
            Bitmap bitmap = Properties.Resources.aps1.ToBitmap();
            IntPtr hBitmap = bitmap.GetHbitmap();
            ImageSource wpfBitmap =
                 Imaging.CreateBitmapSourceFromHBitmap(
                      hBitmap, IntPtr.Zero, Int32Rect.Empty,
                      BitmapSizeOptions.FromEmptyOptions());
            this.Icon = wpfBitmap;
            LoadPSFile.Icon = new System.Windows.Controls.Image
            {
                Source = new BitmapImage(new Uri("Resources/script_go.png", UriKind.Relative))
            };
            AddNewPresetScript.Icon = new System.Windows.Controls.Image
            {
                Source = new BitmapImage(new Uri("Resources/add.png", UriKind.Relative))
            };
            RemovePresetScript.Icon = new System.Windows.Controls.Image
            {
                Source = new BitmapImage(new Uri("Resources/delete.png", UriKind.Relative))
            };

            Title = "Async PowerShell Scripter";
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            LoadingOverlayText.Text = "Loading all preset scripts... Please wait.";
            InitializePresetScriptCombobox();

            Closing += MainWindow_Closing;
            StartButton.Click += StartButton_Click;
            StopButton.Click += StopButton_Click;
            OutputListboxContextMenu.Opened += OutputListboxContextMenu_Opened;
            MenuSelectAll.Click += MenuSelectAll_Click;
            MenuExportSelected.Click += MenuExportSelected_Click;
            MenuExportAll.Click += MenuExportAll_Click;
            ScriptTextbox.AllowDrop = true;
            ScriptTextbox.PreviewDragOver += ScriptTextbox_PreviewDragOver;
            ScriptTextbox.PreviewDrop += ScriptTextbox_PreviewDrop;
            LoadPSFile.Click += LoadPSFile_Click;
            ShowAbout.Click += ShowAbout_Click;
            AddNewPresetScript.Click += AddNewPresetScript_Click;
            RemovePresetScript.Click += RemovePresetScript_Click;

            runSpace = RunspaceFactory.CreateRunspace();
            runSpace.Open();
        }

        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            StopScript();
            runSpace.Close();
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            StopScript();
            OutputListbox.Items.Clear();
            AppendLine("Starting script...");
            StartScript();
        }

        private void StopButton_Click(object sender, RoutedEventArgs e)
        {
            StopScript();
        }

        #region PipelineExecutor
        private void StartScript()
        {
            pipelineExecutor = new PipelineExecutor(runSpace, this.Dispatcher, ScriptTextbox.Text);
            pipelineExecutor.OnDataReady += new PipelineExecutor.DataReadyDelegate(pipelineExecutor_OnDataReady);
            pipelineExecutor.OnDataEnd += new PipelineExecutor.DataEndDelegate(pipelineExecutor_OnDataEnd);
            pipelineExecutor.OnErrorReady += new PipelineExecutor.ErrorReadyDelegate(pipelineExecutor_OnErrorReady);
            pipelineExecutor.Start();
        }

        private void StopScript()
        {
            if (pipelineExecutor != null)
            {
                pipelineExecutor.OnDataReady -= new PipelineExecutor.DataReadyDelegate(pipelineExecutor_OnDataReady);
                pipelineExecutor.OnDataEnd -= new PipelineExecutor.DataEndDelegate(pipelineExecutor_OnDataEnd);
                pipelineExecutor.Stop();
                pipelineExecutor = null;
            }
        }

        private void pipelineExecutor_OnDataEnd(PipelineExecutor sender)
        {
            if (sender.Pipeline.PipelineStateInfo.State == PipelineState.Failed)
            {
                AppendLine(string.Format("Error in script: {0}", sender.Pipeline.PipelineStateInfo.Reason));
            }
            else
            {
                AppendLine("Starting script complete.");
            }
        }

        private void pipelineExecutor_OnDataReady(PipelineExecutor sender, ICollection<PSObject> data)
        {
            foreach (PSObject obj in data)
            {
                AppendLine(obj.ToString());
            }
        }

        private void pipelineExecutor_OnErrorReady(PipelineExecutor sender, ICollection<object> data)
        {
            foreach (object e in data)
            {
                AppendLine("Error : " + e.ToString());
            }
        }
        #endregion PipelineExecutor

        #region PresetScriptCombobox
        private class PresetScript
        {
            public string Name { get; private set; }
            public string Script { get; private set; }
            public PresetScript(string name, string script)
            {
                Name = name;
                Script = script;
            }
        }

        private void InitializePresetScriptCombobox()
        {
            LoadingOverlay.Visibility = Visibility.Visible;
            presetScriptItems = new List<PresetScript>();

            Progressbar.Value = 0;
            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += (obj, e) =>
            {
                try
                {
                    //string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                    Directory.CreateDirectory("presets");
                    string[] fileList = Directory.GetFiles("presets");
                    int count = fileList.Length;

                    string filename;
                    for (int i = 0; i < count; i++)
                    {
                        filename = Path.GetFileNameWithoutExtension(fileList[i]);
                        presetScriptItems.Add(
                            new PresetScript(
                                filename,
                                GetContent(fileList[i])));

                        double percent = ((i + 1) * 100) / count;
                        worker.ReportProgress((int)percent);
                        Thread.Sleep(500);
                    }
                }
                catch
                {
                    throw new Exception();
                }
            };
            worker.ProgressChanged += (obj, e) =>
            {
                //Progressbar.Value = e.ProgressPercentage;
                DoubleAnimation animation = new DoubleAnimation(e.ProgressPercentage, TimeSpan.FromMilliseconds(500));
                Progressbar.BeginAnimation(ProgressBar.ValueProperty, animation);
            };
            worker.RunWorkerCompleted += (obj, e) =>
            {
                PresetScriptCombobox.ItemsSource = presetScriptItems;
                PresetScriptCombobox.DisplayMemberPath = "Name";
                PresetScriptCombobox.SelectedValuePath = "Script";
                PresetScriptCombobox.SelectionChanged += PresetScriptCombobox_SelectionChanged;

                if (savingNewScript)
                {
                    PresetScriptCombobox.SelectedItem = presetScriptItems.FirstOrDefault(x => x.Name == savingNewScriptName);
                    savingNewScript = false;
                }
                else
                {
                    PresetScriptCombobox.SelectedIndex = -1;
                }

                LoadingOverlay.Visibility = Visibility.Collapsed;
            };
            worker.RunWorkerAsync();
        }

        private void PresetScriptCombobox_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (PresetScriptCombobox.SelectedIndex != -1)
                ScriptTextbox.Text = PresetScriptCombobox.SelectedValue.ToString();
        }

        private string GetContent(string filepath)
        {
            using (StreamReader sr = new StreamReader(filepath))
            {
                return sr.ReadToEnd();
            }
        }
        #endregion PresetScriptCombobox

        #region OutputListbox
        private void AppendLine(string line)
        {
            if (OutputListbox.Items.Count > 10000)
                OutputListbox.Items.RemoveAt(0);

            OutputListbox.Items.Add(line);
            OutputListbox.SelectedIndex = OutputListbox.Items.Count - 1;
            OutputListbox.ScrollIntoView(OutputListbox.SelectedItem);
            OutputListbox.UnselectAll();
        }
        #endregion OutputListbox

        #region OutputListboxContextMenu
        private void OutputListboxContextMenu_Opened(object sender, RoutedEventArgs e)
        {
            MenuSelectAll.IsEnabled = true;
            MenuExportSelected.IsEnabled = true;
            MenuExportAll.IsEnabled = true;

            if (OutputListbox.Items.Count == 0)
            {
                MenuSelectAll.IsEnabled = false;
                MenuExportSelected.IsEnabled = false;
                MenuExportAll.IsEnabled = false;
            }
        }

        private void MenuSelectAll_Click(object sender, RoutedEventArgs e)
        {
            OutputListbox.SelectAll();
        }

        private void MenuExportSelected_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = OutputListbox.SelectedItems;

            var saveFile = new SaveFileDialog();
            saveFile.Filter = "Text file (*.txt)|*.txt";
            saveFile.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (saveFile.ShowDialog() == true)
            {
                try
                {
                    using (FileStream fs = File.Open(saveFile.FileName, FileMode.CreateNew))
                    {
                        using (StreamWriter sw = new StreamWriter(fs))
                        {
                            foreach (var item in selectedItems)
                                sw.WriteLine(item.ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        ex.ToString(),
                        "Failed to export",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
                OutputListbox.UnselectAll();
            }
        }

        private void MenuExportAll_Click(object sender, RoutedEventArgs e)
        {
            if (OutputListbox.Items.Count == 0)
                return;

            var saveFile = new SaveFileDialog();
            saveFile.Filter = "Text file (*.txt)|*.txt";
            saveFile.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (saveFile.ShowDialog() == true)
            {
                try
                {
                    using (FileStream fs = File.Open(saveFile.FileName, FileMode.CreateNew))
                    {
                        using (StreamWriter sw = new StreamWriter(fs))
                        {
                            foreach (var item in OutputListbox.Items)
                                sw.WriteLine(item.ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        ex.ToString(),
                        "Failed to export",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
            }
        }
        #endregion OutputListboxContextMenu

        #region ScriptTextbox Drag n Drop
        private void ScriptTextbox_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }

        private void ScriptTextbox_PreviewDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files != null && files.Length != 0)
            {
                if (files.Length > 0)
                {
                    using (StreamReader sr = new StreamReader(files[0]))
                    {
                        ScriptTextbox.Text = sr.ReadToEnd();
                    }
                }
            }

            PresetScriptCombobox.SelectedIndex = -1;
        }
        #endregion ScriptTextbox Drag n Drop

        #region Main Menu
        private void LoadPSFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PowerShell script (*.ps1)|*.ps1";
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (openFileDialog.ShowDialog() == true)
            {
                if (openFileDialog.FileName.Length > 0)
                {
                    using (StreamReader sr = new StreamReader(openFileDialog.FileName))
                    {
                        ScriptTextbox.Text = sr.ReadToEnd();
                    }
                }
            }

            PresetScriptCombobox.SelectedIndex = -1;
        }

        private void AddNewPresetScript_Click(object sender, RoutedEventArgs e)
        {
            Window window = new Window
            {
                Title = "Enter script title",
                Content = new AddPresetScript(this),
                SizeToContent = SizeToContent.WidthAndHeight,
                ResizeMode = ResizeMode.NoResize,
                Owner = this,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Icon = this.Icon,
                ShowInTaskbar = false
            };

            window.ShowDialog();
        }

        public void SaveNewScript(string title)
        {
            try
            {
                string filename = MakeValidFileName(title);
                Directory.CreateDirectory("presets");
                string baseDir = AppDomain.CurrentDomain.BaseDirectory + "presets";
                string fullpath = baseDir + "\\" + filename + ".txt";
                //if (File.Exists(fullpath))
                //    fullpath = baseDir + "\\" + MakeValidFileName(title) + " - Copy.txt";

                File.WriteAllText(fullpath, ScriptTextbox.Text);

                savingNewScript = true;
                savingNewScriptName = filename;
                ScriptTextbox.Text = "";
                PresetScriptCombobox.SelectionChanged -= PresetScriptCombobox_SelectionChanged;
                PresetScriptCombobox.ItemsSource = null;
                LoadingOverlayText.Text = "Reloading all preset scripts... Please wait.";
                InitializePresetScriptCombobox();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.ToString(),
                    "Failed to save",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void RemovePresetScript_Click(object sender, RoutedEventArgs e)
        {
            if (PresetScriptCombobox.SelectedItem == null)
                return;

            try
            {
                var selectedPreset = PresetScriptCombobox.SelectedItem as PresetScript;
                string filename = selectedPreset.Name;
                string baseDir = AppDomain.CurrentDomain.BaseDirectory + "presets";
                string fullpath = baseDir + "\\" + filename + ".txt";
                if (File.Exists(fullpath))
                {
                    File.Delete(fullpath);
                }

                ScriptTextbox.Text = "";
                PresetScriptCombobox.SelectionChanged -= PresetScriptCombobox_SelectionChanged;
                PresetScriptCombobox.ItemsSource = null;
                LoadingOverlayText.Text = "Reloading all preset scripts... Please wait.";
                InitializePresetScriptCombobox();
            }
            catch
            {
                return;
            }
        }

        private void ShowAbout_Click(object sender, RoutedEventArgs e)
        {
            var version = AssemblyName.GetAssemblyName(Assembly.GetExecutingAssembly().Location).Version.ToSt‌​ring(2);
            MessageBox.Show(
                string.Format("{0} {1}\n\nCreated by Heiswayi Nrird\nWebsite: https://heiswayi.github.io\n\nIt's freeware!", Title, version),
                "About this app",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        #endregion Main Menu

        #region Helper functions
        private static string MakeValidFileName(string name)
        {
            string invalidChars = System.Text.RegularExpressions.Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
            string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);

            return System.Text.RegularExpressions.Regex.Replace(name, invalidRegStr, "_");
        }
        #endregion Helper functions
    }
}