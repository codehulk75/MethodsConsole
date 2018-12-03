using ModernChrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Text.RegularExpressions;
using System.IO;
using System.Reflection;
using System.ComponentModel;

namespace Methods_Console
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 


    public partial class MainWindow
    { 
        int themenum = Properties.Settings.Default.WindowThemeNumber;
        int borderbrushnum = Properties.Settings.Default.StatusBarThemeNumber;
        Dictionary<int, string> themes = new Dictionary<int, string>();
        Dictionary<int, string> borderBrushes = new Dictionary<int, string>();
        List<Ci2Parser> ExportList = new List<Ci2Parser>();
        BeiBOM Bom;
        Assembly thisProgram;
        AssemblyName thisProgramName;
        Version thisProgramVersion;
        List<TextBox> lstProgTextBoxes = new List<TextBox>();
        List<ComboBox> lstProgComboBoxes = new List<ComboBox>();
        List<string> lstPasses = new List<string>(new string[] { "SMT 1", "SMT 2" });
        List<ComboBox> lstMachComboBoxes = new List<ComboBox>();
        List<Label> lstProgDateLabels = new List<Label>();
        List<Button> lstAddProgramButtons = new List<Button>();
        List<Button> lstRemoveProgramButtons = new List<Button>();
        SortedDictionary<int, string> dictMachines = new SortedDictionary<int, string>();
        public List<string> MachineNames { get; private set; }
        public MainWindow()
        {
            thisProgram = Assembly.GetEntryAssembly();
            thisProgramName = thisProgram.GetName();
            thisProgramVersion = thisProgramName.Version;
            InitializeComponent();
            MachineNames = Properties.Settings.Default.SmtMachines.Split(';').ToList();
            System.IO.Directory.CreateDirectory(@"C:\BaaN-DAT");
            InitializeThemeColors();
            try
            {
                ThemeManager.ChangeTheme(Application.Current, themes[themenum]);
                BorderBrush = Application.Current.FindResource($"StatusBar{borderBrushes[borderbrushnum]}BrushKey") as SolidColorBrush;
                SetBrushes(themes[themenum]);
            }
            catch (Exception)
            {
                MessageBox.Show("Error setting window skin.\nMainWindow()->ThemeManager.ChangeTheme())");
            }
            ThemeManager.ThemeChanged += (sender, args) =>
            {
                SetTheme(args.Theme, borderbrushnum);
            };
            int mnum = 0;
            foreach(string machine in MachineNames)
            {
                dictMachines.Add(mnum, machine);
                ++mnum;
            }
            textBoxBom.IsReadOnly = true;
            textBoxBomRev.IsReadOnly = true;
            textBoxBomDate.IsReadOnly = true;
            textBoxBom.Visibility = Visibility.Hidden;
            textBoxBomRev.Visibility = Visibility.Hidden;
            textBoxBomDate.Visibility = Visibility.Hidden;
            labelLoadingOne.Visibility = Visibility.Hidden;
            labelLoadingTwo.Visibility = Visibility.Hidden;
            tbLoadingInsOne.Visibility = Visibility.Hidden;
            tbLoadingInsTwo.Visibility = Visibility.Hidden;
            btnCreate.Visibility = Visibility.Hidden;
            btnFlipLine.Visibility = Visibility.Hidden;
            lstProgTextBoxes.Add(textBoxProg1);
            lstProgTextBoxes.Add(textBoxProg2);
            lstProgTextBoxes.Add(textBoxProg3);
            lstProgTextBoxes.Add(textBoxProg4);
            lstProgTextBoxes.Add(textBoxProg5);
            lstProgTextBoxes.Add(textBoxProg6);
            lstProgTextBoxes.Add(textBoxProg7);
            lstProgTextBoxes.Add(textBoxProg8);
            lstProgComboBoxes.Add(comboBoxProg1);
            lstProgComboBoxes.Add(comboBoxProg2);
            lstProgComboBoxes.Add(comboBoxProg3);
            lstProgComboBoxes.Add(comboBoxProg4);
            lstProgComboBoxes.Add(comboBoxProg5);
            lstProgComboBoxes.Add(comboBoxProg6);
            lstProgComboBoxes.Add(comboBoxProg7);
            lstProgComboBoxes.Add(comboBoxProg8);
            lstMachComboBoxes.Add(comboBoxMach1);
            lstMachComboBoxes.Add(comboBoxMach2);
            lstMachComboBoxes.Add(comboBoxMach3);
            lstMachComboBoxes.Add(comboBoxMach4);
            lstMachComboBoxes.Add(comboBoxMach5);
            lstMachComboBoxes.Add(comboBoxMach6);
            lstMachComboBoxes.Add(comboBoxMach7);
            lstMachComboBoxes.Add(comboBoxMach8);
            lstProgDateLabels.Add(labelProg1);
            lstProgDateLabels.Add(labelProg2);
            lstProgDateLabels.Add(labelProg3);
            lstProgDateLabels.Add(labelProg4);
            lstProgDateLabels.Add(labelProg5);
            lstProgDateLabels.Add(labelProg6);
            lstProgDateLabels.Add(labelProg7);
            lstProgDateLabels.Add(labelProg8);
            lstAddProgramButtons.Add(btnAdd1);
            lstAddProgramButtons.Add(btnAdd2);
            lstAddProgramButtons.Add(btnAdd3);
            lstAddProgramButtons.Add(btnAdd4);
            lstAddProgramButtons.Add(btnAdd5);
            lstAddProgramButtons.Add(btnAdd6);
            lstAddProgramButtons.Add(btnAdd7);
            lstRemoveProgramButtons.Add(btnRemoveProgram1);
            lstRemoveProgramButtons.Add(btnRemoveProgram2);
            lstRemoveProgramButtons.Add(btnRemoveProgram3);
            lstRemoveProgramButtons.Add(btnRemoveProgram4);
            lstRemoveProgramButtons.Add(btnRemoveProgram5);
            lstRemoveProgramButtons.Add(btnRemoveProgram6);
            lstRemoveProgramButtons.Add(btnRemoveProgram7);
            
            foreach(Button btn in lstRemoveProgramButtons)
            {
                btn.Visibility = Visibility.Hidden;
            }

            foreach (TextBox box in lstProgTextBoxes)
            {
                box.Visibility = Visibility.Hidden;
                box.IsReadOnly = true;
            }
            foreach(ComboBox cb in lstProgComboBoxes)
            {
                cb.Visibility = Visibility.Hidden;
                cb.Items.Add(lstPasses[0]);
                cb.Items.Add(lstPasses[1]);
                cb.SelectedIndex = 0;
            }
            foreach(ComboBox cb in lstMachComboBoxes)
            {
                cb.Visibility = Visibility.Hidden;
                foreach(var machine in dictMachines)
                {
                    cb.Items.Add(machine.Value);
                }
                cb.SelectedIndex = 0;
            }
            foreach (Label lb in lstProgDateLabels)
            {
                lb.Visibility = Visibility.Hidden;
            }
            foreach(Button bt in lstAddProgramButtons)
            {
                bt.Visibility = Visibility.Hidden;
            }

        }
        private bool IsDoublePassAssembly()
        {
            bool bHasTwoPasses = false;
            HashSet<string> hsSides = new HashSet<string>();            
            foreach(var export in ExportList)
            {
                string side = export.FileName.Contains("BOT") ? "SMT 1" : "SMT 2";
                hsSides.Add(side);
            }
            if (hsSides.Count > 1)
                bHasTwoPasses = true;
            return bHasTwoPasses;
        }
        private void ClearData()
        {
            foreach(TextBox tb in lstProgTextBoxes)
            {
                tb.Clear();
                tb.Visibility = Visibility.Hidden;
            }
            foreach (ComboBox cb in lstMachComboBoxes)
            {
                cb.Visibility = Visibility.Hidden;
            }
            foreach (ComboBox cb in lstProgComboBoxes)
            {
                cb.Visibility = Visibility.Hidden;
            }
            foreach (Label lbl in lstProgDateLabels)
            {
                lbl.Content = "";
                lbl.Visibility = Visibility.Hidden;
            }
            foreach(Button btn in lstAddProgramButtons)
            {
                btn.Visibility = Visibility.Hidden;
            }
            foreach (Button btn in lstRemoveProgramButtons)
            {
                btn.Visibility = Visibility.Hidden;
            }
            btnCreate.Visibility = Visibility.Hidden;
            btnFlipLine.Visibility = Visibility.Hidden;
            textBoxBom.Clear();
            textBoxBomRev.Clear();
            textBoxBomDate.Clear();
            tbLoadingInsOne.Clear();
            tbLoadingInsTwo.Clear();
            textBoxBom.Visibility = Visibility.Hidden;
            textBoxBomRev.Visibility = Visibility.Hidden;
            textBoxBomDate.Visibility = Visibility.Hidden;
            labelLoadingOne.Visibility = Visibility.Hidden;
            labelLoadingTwo.Visibility = Visibility.Hidden;
            tbLoadingInsOne.Visibility = Visibility.Hidden;
            tbLoadingInsTwo.Visibility = Visibility.Hidden;
            ExportList.Clear();
            Bom = null;
        }

        private void InitializeThemeColors()
        {
            themes.Add(1, "LightBlue");
            themes.Add(2, "DarkBlue");
            themes.Add(3, "Light");
            themes.Add(4, "Dark");
            themes.Add(5, "Blend");
            borderBrushes.Add(1, "Blue");
            borderBrushes.Add(2, "Orange");
            borderBrushes.Add(3, "Purple");
            borderBrushes.Add(4, "Green");
        }

        private void SetBrushes(String strThemeColor)
        {
            if (strThemeColor != "Light")
            {
                TasksText.Foreground = Brushes.White;
                setupsheetbutton.Foreground = Brushes.White;
                famcheckbutton.Foreground = Brushes.White;
                themebutton.Foreground = Brushes.White;
                comparebutton.Foreground = Brushes.White;
            }
            else
            {
                TasksText.Foreground = Brushes.Black;
                setupsheetbutton.Foreground = Brushes.Black;
                famcheckbutton.Foreground = Brushes.Black;
                themebutton.Foreground = Brushes.Black;
                comparebutton.Foreground = Brushes.Black;
            }
        }

        private void SetTheme(String strThemeColor, int nBorderColor)
        {
            BorderBrush = Application.Current.FindResource($"StatusBar{borderBrushes[nBorderColor]}BrushKey") as SolidColorBrush;
            SetBrushes(strThemeColor);
            int colorKey = themes.FirstOrDefault(x => x.Value == strThemeColor).Key;
            Properties.Settings.Default.StatusBarThemeNumber = nBorderColor;
            Properties.Settings.Default.WindowThemeNumber = colorKey;
            Properties.Settings.Default.Save();
        }

        private void button_theme_Click(object sender, RoutedEventArgs e)
        {
            ++themenum;
            if (themenum > 5)
                themenum = 1;
            ThemeManager.ChangeTheme(Application.Current, themes[themenum]);
        }

        private void themebutton_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            ++borderbrushnum;
            if (borderbrushnum > 4)
                borderbrushnum = 1;
            SetTheme(themes[themenum], borderbrushnum);
        }

        private void LoadBOM()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xls";
            dlg.Filter = "BOM Explosion/BAAN BOM|*.xls;*xlsx;*.csv;*.txt";
            dlg.Multiselect = false;
            dlg.Title = "Load the BOM File";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                if (Bom != null)
                    Bom = null;
                Bom = new BeiBOM(dlg.FileName);
                if (Bom.IsValid == false)
                    MessageBox.Show("Invalid BOM");
                textBoxBom.Text = "Assembly: " + Bom.AssemblyName;
                textBoxBomRev.Text = Bom.Rev;
                textBoxBomDate.Text = Bom.DateOfListing;
                textBoxBom.Visibility = Visibility.Visible;
                textBoxBomRev.Visibility = Visibility.Visible;
                textBoxBomDate.Visibility = Visibility.Visible;
            }
            else
            {
                Bom = null;
                textBoxBom.Text = "Assembly: ";
                textBoxBomRev.Text = "";
                textBoxBomDate.Text = "";
                textBoxBom.Visibility = Visibility.Visible;
                textBoxBomRev.Visibility = Visibility.Visible;
                textBoxBomDate.Visibility = Visibility.Visible;
            }
        }

        private void PrepareSetupSheet()
        {          
            Tuple<bool, List<string>, List<string>> tDetectInfo;
            tDetectInfo = AutoDetectExportInfo();
            if (tDetectInfo.Item1 == false)
            {
                string strFails = "Failed Pass Detection for the following programs:\n";
                foreach(string s in tDetectInfo.Item2)
                {
                    strFails += s + "\n";
                }
                strFails += "\nFailed Machine Detection for the following programs:\n";
                foreach(string s in tDetectInfo.Item3)
                {
                    strFails += s + "\n";
                }
                strFails += "\nPlease manually choose pass/machine for these.";       
                MessageBox.Show(strFails, "Could Not Auto Detect", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            ExportList = ExportList.OrderBy(p => p.Pass).ThenBy(m => m.MachineName).ToList();
            PopulateUIBoxes();
        }

        private void PopulateUIBoxes()
        {
            int i = 0;
            bool bHasTwoPasses = false;
            if(Bom != null)
                btnCreate.Visibility = Visibility.Visible;
            labelLoadingOne.Visibility = Visibility.Visible;
            tbLoadingInsOne.Visibility = Visibility.Visible;

            foreach( Ci2Parser program in ExportList)
            {
                lstProgTextBoxes[i].Text = program.ProgramName;
                lstProgTextBoxes[i].Visibility = Visibility.Visible;
                lstProgComboBoxes[i].Visibility = Visibility.Visible;
                lstMachComboBoxes[i].Visibility = Visibility.Visible;
                lstProgDateLabels[i].Content = program.DateCreated;
                lstProgDateLabels[i].Visibility = Visibility.Visible;              
                if (program.Pass.Equals(lstPasses[0]))
                { 
                    lstProgComboBoxes[i].SelectedIndex = 0;
                }
                else if (program.Pass.Equals(lstPasses[1]))
                {
                    if (IsDoublePassAssembly())
                    {
                        lstProgComboBoxes[i].SelectedIndex = 1;
                        bHasTwoPasses = true;
                    }
                    else
                        lstProgComboBoxes[i].SelectedIndex = 0;
                }
                foreach(var item in dictMachines)
                {
                    if (program.MachineName.Equals(item.Value))
                        lstMachComboBoxes[i].SelectedIndex = item.Key;
                }
                ++i;
            }
            if(bHasTwoPasses == true)
            {
                btnFlipLine.Visibility = Visibility.Visible;
                tbLoadingInsTwo.Visibility = Visibility.Visible;
                labelLoadingTwo.Visibility = Visibility.Visible;
            }
            if (ExportList.Count < 8 && ExportList.Count > 0)
                lstAddProgramButtons[ExportList.Count - 1].Visibility = Visibility.Visible;
            if (ExportList.Count < 9 && ExportList.Count > 1)
                lstRemoveProgramButtons[ExportList.Count - 2].Visibility = Visibility.Visible;
        }
        private void SwapUIBoxes()
        {
            int i = 0;
            string strCurrentPass = lstProgComboBoxes.First().Text;
            if (strCurrentPass.Equals("SMT 1"))
            {
                ExportList = ExportList.OrderByDescending(p => p.Pass).ThenBy(m => m.MachineName).ToList();              
            }
            else if (strCurrentPass.Equals("SMT 2"))
            {
                ExportList = ExportList.OrderBy(p => p.Pass).ThenBy(m => m.MachineName).ToList();
            }
            foreach(var export in ExportList)
            {
                export.Pass = (string)lstProgComboBoxes[i].SelectedValue;
                lstProgTextBoxes[i].Text = export.ProgramName;
                lstProgDateLabels[i].Content = export.DateCreated;
                foreach (var item in dictMachines)
                {
                    if (export.MachineName.Equals(item.Value))
                        lstMachComboBoxes[i].SelectedIndex = item.Key;
                }
                ++i;
            }
            if(tbLoadingInsOne.Visibility == Visibility.Visible && tbLoadingInsTwo.Visibility == Visibility.Visible)
            {
                string t = tbLoadingInsOne.Text;
                tbLoadingInsOne.Text = tbLoadingInsTwo.Text;
                tbLoadingInsTwo.Text = t;
            }
        }
        private Tuple<bool, List<string>, List<string>> AutoDetectExportInfo(bool bDescending = false)
        {
            //
            //Try to auto detect Pass and machine name based on ci2 file name
            //
            bool bDectectionSuccessful = false;
            List<string> failedMachineDetection = new List<string>();
            List<string> failedPassDetection = new List<string>();
            Regex reBot = new Regex(@"\bBOT\b", RegexOptions.IgnoreCase);
            Regex reTop = new Regex(@"\bTOP\b", RegexOptions.IgnoreCase);
            Regex reFuzion = new Regex(@"\s+FZ60XC\s+",RegexOptions.IgnoreCase);
            Regex reGCOne = new Regex(@"\s+(GC-?_?60-?_?1?)\s+", RegexOptions.IgnoreCase);
            Regex reGCTwo = new Regex(@"\s+(GC-?_?60-?_?2)\s+", RegexOptions.IgnoreCase);
            Regex reGI = new Regex(@"\s+(GI-?_?14-?_?\d?)\s+", RegexOptions.IgnoreCase);
            Regex reGX = new Regex(@"\s+(GX-?_?11)\s+", RegexOptions.IgnoreCase);

            foreach (Ci2Parser parser in ExportList)
            {
                string filename = parser.ProgramName;
                if (reBot.IsMatch(filename))
                {
                    if (bDescending == false || IsDoublePassAssembly() == false)
                        parser.Pass = lstPasses[0];
                    else
                        parser.Pass = lstPasses[1];
                }
                else if (reTop.IsMatch(filename))
                {
                    if (bDescending == false && IsDoublePassAssembly())
                        parser.Pass = lstPasses[1];
                    else
                        parser.Pass = lstPasses[0];
                }
                else
                {
                    failedPassDetection.Add(parser.ProgramName);
                }
                if (reFuzion.IsMatch(filename))
                {
                    parser.MachineName = dictMachines[0];
                }
                else if (reGCOne.IsMatch(filename))
                {
                    parser.MachineName = dictMachines[1];
                }
                else if (reGCTwo.IsMatch(filename))
                {
                    parser.MachineName = dictMachines[2];
                }
                else if (reGI.IsMatch(filename))
                {
                    parser.MachineName = dictMachines[3];
                }
                else if (reGX.IsMatch(filename))
                {
                    parser.MachineName = dictMachines[4];
                }
                else
                {
                    bool bFound = false;
                    foreach (string machine in MachineNames)
                    {
                        if (filename.Contains(machine))
                        {
                            parser.MachineName = machine;
                            bFound = true;
                        }
                    }
                    if(bFound == false)
                        failedMachineDetection.Add(parser.ProgramName);
                }
            }
            if (failedMachineDetection.Count == 0 && failedPassDetection.Count == 0)
                bDectectionSuccessful = true;

            return Tuple.Create(bDectectionSuccessful, failedPassDetection, failedMachineDetection);

        }

        private void setupsheetbutton_Click(object sender, RoutedEventArgs e)
        {
            
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".ci2";
            dlg.Filter = "Universal Export|*.ci2|All Files|*.*";
            dlg.Multiselect = true;
            dlg.Title = "Load the Program Exports";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                ClearData();
                string[] files = dlg.FileNames;
                if(files.Length > 8)
                {
                    MessageBox.Show("Too many files.  At this time the maximum number you can choose is 8.", "Exceeded Export File Limit", MessageBoxButton.OK, MessageBoxImage.Hand);
                    return;
                }
                foreach (var filename in files)
                {
                    FileParserFactory factory = new Ci2ParserFactory(filename);
                    Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                    if(parser != null)
                    {
                        ExportList.Add(parser);
                    }
                }
                LoadBOM();
                PrepareSetupSheet();
            }
        }

        private void btnCreate_Click(object sender, RoutedEventArgs e)
        {
            (sender as Button).IsEnabled = false;
            ProgressWindow pw = new ProgressWindow(@"C:\BaaN-DAT\"+Bom.AssemblyName + "_" + Bom.Rev + ".rtf");
            pw.Owner = this;
            pw.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            try
            {
                string filename = @"C:\BaaN-DAT\" + Bom.AssemblyName + "_" + Bom.Rev + ".rtf";
                using (StreamWriter writer = new StreamWriter(filename, false)) { }
            }
            catch(IOException ioException)
            {
                (sender as Button).IsEnabled = true;
                MessageBox.Show("Cannot create setup sheet file. If you have it open please close it and try again.\n" + ioException.Message, "Error Creating Setup Sheet", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            List<string> lstLoadingInstructions = new List<string>(new string[] { tbLoadingInsOne.Text, tbLoadingInsTwo.Text });

            BackgroundWorker worker = new BackgroundWorker();
            worker.DoWork += worker_DoWork;
            worker.RunWorkerCompleted += (object s, RunWorkerCompletedEventArgs eargs) => {
                if (pw != null && pw.Visibility == Visibility.Visible)
                {
                    pw.Title = "Done Processing";
                    pw.progressBar.Visibility = Visibility.Hidden;
                    pw.StatusLabel.Content = "Setup Sheet Complete!";
                    pw.OpenButton.Visibility = Visibility.Visible;
                    pw.editButton.Visibility = Visibility.Visible;
                    pw.closeButton.Visibility = Visibility.Visible;
                    (sender as Button).IsEnabled = true;
                } };
            pw.Show();
            worker.RunWorkerAsync(lstLoadingInstructions);     
        }
        void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            SetupSheetGenerator setupSheet = new SetupSheetGenerator(Bom, ExportList, (List<string>)e.Argument);
            setupSheet.CreateSheet();
        }

        private void comboBoxMach1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 0)
                ExportList[0].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 1)
                ExportList[1].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 2)
                ExportList[2].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach4_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 3)
                ExportList[3].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach5_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 4)
                ExportList[4].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach6_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 5)
                ExportList[5].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach7_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 6)
                ExportList[6].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxMach8_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 7)
                ExportList[7].MachineName = (sender as ComboBox).SelectedItem.ToString();
        }

        private void comboBoxProg1_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 0)
                ExportList[0].Pass = (sender as ComboBox).SelectedItem.ToString();
            bool bTwoPasses = false;
            foreach (var cb in lstProgComboBoxes)
            {
                if (cb.Visibility == Visibility.Visible && cb.SelectedIndex == 1)
                    bTwoPasses = true;
            }
            if (bTwoPasses)
            {
                labelLoadingTwo.Visibility = Visibility.Visible;
                tbLoadingInsTwo.Visibility = Visibility.Visible;
            }
            else
            {
                labelLoadingTwo.Visibility = Visibility.Hidden;
                tbLoadingInsTwo.Visibility = Visibility.Hidden;
            }
        }

        private void comboBoxProg2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 1)
                ExportList[1].Pass = (sender as ComboBox).SelectedItem.ToString();
            bool bTwoPasses = false;
            foreach (var cb in lstProgComboBoxes)
            {
                if (cb.Visibility == Visibility.Visible && cb.SelectedIndex == 1)
                    bTwoPasses = true;
            }
            if (bTwoPasses)
            {
                labelLoadingTwo.Visibility = Visibility.Visible;
                tbLoadingInsTwo.Visibility = Visibility.Visible;
            }
            else
            {
                labelLoadingTwo.Visibility = Visibility.Hidden;
                tbLoadingInsTwo.Visibility = Visibility.Hidden;
            }
        }

        private void comboBoxProg3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 2)
                ExportList[2].Pass = (sender as ComboBox).SelectedItem.ToString();
            bool bTwoPasses = false;
            foreach (var cb in lstProgComboBoxes)
            {
                if (cb.Visibility == Visibility.Visible && cb.SelectedIndex == 1)
                    bTwoPasses = true;
            }
            if (bTwoPasses)
            {
                labelLoadingTwo.Visibility = Visibility.Visible;
                tbLoadingInsTwo.Visibility = Visibility.Visible;
            }
            else
            {
                labelLoadingTwo.Visibility = Visibility.Hidden;
                tbLoadingInsTwo.Visibility = Visibility.Hidden;
            }
        }

        private void comboBoxProg4_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 3)
                ExportList[3].Pass = (sender as ComboBox).SelectedItem.ToString();
            bool bTwoPasses = false;
            foreach (var cb in lstProgComboBoxes)
            {
                if (cb.Visibility == Visibility.Visible && cb.SelectedIndex == 1)
                    bTwoPasses = true;
            }
            if (bTwoPasses)
            {
                labelLoadingTwo.Visibility = Visibility.Visible;
                tbLoadingInsTwo.Visibility = Visibility.Visible;
            }
            else
            {
                labelLoadingTwo.Visibility = Visibility.Hidden;
                tbLoadingInsTwo.Visibility = Visibility.Hidden;
            }
        }

        private void comboBoxProg5_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 4)
                ExportList[4].Pass = (sender as ComboBox).SelectedItem.ToString();
            bool bTwoPasses = false;
            foreach (var cb in lstProgComboBoxes)
            {
                if (cb.Visibility == Visibility.Visible && cb.SelectedIndex == 1)
                    bTwoPasses = true;
            }
            if (bTwoPasses)
            {
                labelLoadingTwo.Visibility = Visibility.Visible;
                tbLoadingInsTwo.Visibility = Visibility.Visible;
            }
            else
            {
                labelLoadingTwo.Visibility = Visibility.Hidden;
                tbLoadingInsTwo.Visibility = Visibility.Hidden;
            }
        }

        private void comboBoxProg6_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 5)
                ExportList[5].Pass = (sender as ComboBox).SelectedItem.ToString();
            bool bTwoPasses = false;
            foreach (var cb in lstProgComboBoxes)
            {
                if (cb.Visibility == Visibility.Visible && cb.SelectedIndex == 1)
                    bTwoPasses = true;
            }
            if (bTwoPasses)
            {
                labelLoadingTwo.Visibility = Visibility.Visible;
                tbLoadingInsTwo.Visibility = Visibility.Visible;
            }
            else
            {
                labelLoadingTwo.Visibility = Visibility.Hidden;
                tbLoadingInsTwo.Visibility = Visibility.Hidden;
            }
        }

        private void comboBoxProg7_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 6)
                ExportList[6].Pass = (sender as ComboBox).SelectedItem.ToString();
            bool bTwoPasses = false;
            foreach (var cb in lstProgComboBoxes)
            {
                if (cb.Visibility == Visibility.Visible && cb.SelectedIndex == 1)
                    bTwoPasses = true;
            }
            if (bTwoPasses)
            {
                labelLoadingTwo.Visibility = Visibility.Visible;
                tbLoadingInsTwo.Visibility = Visibility.Visible;
            }
            else
            {
                labelLoadingTwo.Visibility = Visibility.Hidden;
                tbLoadingInsTwo.Visibility = Visibility.Hidden;
            }
        }

        private void comboBoxProg8_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if ((sender as ComboBox).Visibility == Visibility.Visible && ExportList.Count > 7)
                ExportList[7].Pass = (sender as ComboBox).SelectedItem.ToString();
            bool bTwoPasses = false;
            foreach (var cb in lstProgComboBoxes)
            {
                if (cb.Visibility == Visibility.Visible && cb.SelectedIndex == 1)
                    bTwoPasses = true;
            }
            if (bTwoPasses)
            {
                labelLoadingTwo.Visibility = Visibility.Visible;
                tbLoadingInsTwo.Visibility = Visibility.Visible;
            }
            else
            {
                labelLoadingTwo.Visibility = Visibility.Hidden;
                tbLoadingInsTwo.Visibility = Visibility.Hidden;
            }
        }

        private void btnAdd1_Click(object sender, RoutedEventArgs e)
        {
            Button thisButton = sender as Button;
            lstProgTextBoxes[1].Visibility = Visibility.Visible;
            lstProgDateLabels[1].Visibility = Visibility.Visible;
            lstProgComboBoxes[1].Visibility = Visibility.Visible;
            lstMachComboBoxes[1].Visibility = Visibility.Visible;
            lstMachComboBoxes[1].SelectedIndex = 0;
            thisButton.Visibility = Visibility.Hidden;
            btnAdd2.Visibility = Visibility.Visible;
            btnRemoveProgram1.Visibility = Visibility.Visible;
        }

        private void btnAdd2_Click(object sender, RoutedEventArgs e)
        {
            Button thisButton = sender as Button;
            lstProgTextBoxes[2].Visibility = Visibility.Visible;
            lstProgDateLabels[2].Visibility = Visibility.Visible;
            lstProgComboBoxes[2].Visibility = Visibility.Visible;
            lstMachComboBoxes[2].Visibility = Visibility.Visible;
            lstMachComboBoxes[2].SelectedIndex = 0;
            thisButton.Visibility = Visibility.Hidden;
            btnAdd3.Visibility = Visibility.Visible;
            btnRemoveProgram2.Visibility = Visibility.Visible;
            btnRemoveProgram1.Visibility = Visibility.Hidden;
        }

        private void btnAdd3_Click(object sender, RoutedEventArgs e)
        {
            Button thisButton = sender as Button;
            lstProgTextBoxes[3].Visibility = Visibility.Visible;
            lstProgDateLabels[3].Visibility = Visibility.Visible;
            lstProgComboBoxes[3].Visibility = Visibility.Visible;
            lstProgComboBoxes[3].SelectedIndex = 1;
            lstMachComboBoxes[3].Visibility = Visibility.Visible;
            lstMachComboBoxes[3].SelectedIndex = 0;
            thisButton.Visibility = Visibility.Hidden;
            btnAdd4.Visibility = Visibility.Visible;
            btnRemoveProgram3.Visibility = Visibility.Visible;
            btnRemoveProgram2.Visibility = Visibility.Hidden;
        }

        private void btnAdd4_Click(object sender, RoutedEventArgs e)
        {
            Button thisButton = sender as Button;
            lstProgTextBoxes[4].Visibility = Visibility.Visible;
            lstProgDateLabels[4].Visibility = Visibility.Visible;
            lstProgComboBoxes[4].Visibility = Visibility.Visible;
            lstProgComboBoxes[4].SelectedIndex = 1;
            lstMachComboBoxes[4].Visibility = Visibility.Visible;
            lstMachComboBoxes[4].SelectedIndex = 0;
            thisButton.Visibility = Visibility.Hidden;
            btnAdd5.Visibility = Visibility.Visible;
            btnRemoveProgram4.Visibility = Visibility.Visible;
            btnRemoveProgram3.Visibility = Visibility.Hidden;
        }

        private void btnAdd5_Click(object sender, RoutedEventArgs e)
        {
            Button thisButton = sender as Button;
            lstProgTextBoxes[5].Visibility = Visibility.Visible;
            lstProgDateLabels[5].Visibility = Visibility.Visible;
            lstProgComboBoxes[5].Visibility = Visibility.Visible;
            lstProgComboBoxes[5].SelectedIndex = 1;
            lstMachComboBoxes[5].Visibility = Visibility.Visible;
            lstMachComboBoxes[5].SelectedIndex = 0;
            thisButton.Visibility = Visibility.Hidden;
            btnAdd6.Visibility = Visibility.Visible;
            btnRemoveProgram5.Visibility = Visibility.Visible;
            btnRemoveProgram4.Visibility = Visibility.Hidden;
        }

        private void btnAdd6_Click(object sender, RoutedEventArgs e)
        {
            Button thisButton = sender as Button;
            lstProgTextBoxes[6].Visibility = Visibility.Visible;
            lstProgDateLabels[6].Visibility = Visibility.Visible;
            lstProgComboBoxes[6].Visibility = Visibility.Visible;
            lstProgComboBoxes[6].SelectedIndex = 1;
            lstMachComboBoxes[6].Visibility = Visibility.Visible;
            lstMachComboBoxes[6].SelectedIndex = 0;
            thisButton.Visibility = Visibility.Hidden;
            btnAdd7.Visibility = Visibility.Visible;
            btnRemoveProgram6.Visibility = Visibility.Visible;
            btnRemoveProgram5.Visibility = Visibility.Hidden;
        }

        private void btnAdd7_Click(object sender, RoutedEventArgs e)
        {
            Button thisButton = sender as Button;
            lstProgTextBoxes[7].Visibility = Visibility.Visible;
            lstProgDateLabels[7].Visibility = Visibility.Visible;
            lstProgComboBoxes[7].Visibility = Visibility.Visible;
            lstProgComboBoxes[7].SelectedIndex = 1;
            lstMachComboBoxes[7].Visibility = Visibility.Visible;
            lstMachComboBoxes[7].SelectedIndex = 0;
            thisButton.Visibility = Visibility.Hidden;
            btnRemoveProgram7.Visibility = Visibility.Visible;
            btnRemoveProgram6.Visibility = Visibility.Hidden;
        }

        private void btnRemoveProgram1_Click(object sender, RoutedEventArgs e)
        {
            if(ExportList.Count > 1)
                ExportList.RemoveAt(1);
            lstMachComboBoxes[1].Text = "";
            lstMachComboBoxes[1].Visibility = Visibility.Hidden;
            lstProgComboBoxes[1].SelectedIndex = 0;
            lstProgComboBoxes[1].Visibility = Visibility.Hidden;
            lstProgDateLabels[1].Content = "";
            lstProgDateLabels[1].Visibility = Visibility.Hidden;
            lstProgTextBoxes[1].Text = "";
            lstProgTextBoxes[1].Visibility = Visibility.Hidden;
            btnRemoveProgram1.Visibility = Visibility.Hidden;
            btnAdd1.Visibility = Visibility.Visible;
            btnAdd2.Visibility = Visibility.Hidden;
        }

        private void btnRemoveProgram2_Click(object sender, RoutedEventArgs e)
        {
            if (ExportList.Count > 2)
                ExportList.RemoveAt(2);
            lstMachComboBoxes[2].Text = "";
            lstMachComboBoxes[2].Visibility = Visibility.Hidden;
            lstProgComboBoxes[2].SelectedIndex = 0;
            lstProgComboBoxes[2].Visibility = Visibility.Hidden;
            lstProgDateLabels[2].Content = "";
            lstProgDateLabels[2].Visibility = Visibility.Hidden;
            lstProgTextBoxes[2].Text = "";
            lstProgTextBoxes[2].Visibility = Visibility.Hidden;
            btnRemoveProgram2.Visibility = Visibility.Hidden;
            btnAdd2.Visibility = Visibility.Visible;
            btnAdd3.Visibility = Visibility.Hidden;
            btnRemoveProgram1.Visibility = Visibility.Visible;
        }

        private void btnRemoveProgram3_Click(object sender, RoutedEventArgs e)
        {
            if (ExportList.Count > 3)
                ExportList.RemoveAt(3);
            lstMachComboBoxes[3].Text = "";
            lstMachComboBoxes[3].Visibility = Visibility.Hidden;
            lstProgComboBoxes[3].SelectedIndex = 0;
            lstProgComboBoxes[3].Visibility = Visibility.Hidden;
            lstProgDateLabels[3].Content = "";
            lstProgDateLabels[3].Visibility = Visibility.Hidden;
            lstProgTextBoxes[3].Text = "";
            lstProgTextBoxes[3].Visibility = Visibility.Hidden;
            btnRemoveProgram3.Visibility = Visibility.Hidden;
            btnAdd3.Visibility = Visibility.Visible;
            btnAdd4.Visibility = Visibility.Hidden;
            btnRemoveProgram2.Visibility = Visibility.Visible;
        }

        private void btnRemoveProgram4_Click(object sender, RoutedEventArgs e)
        {
            if (ExportList.Count > 4)
                ExportList.RemoveAt(4);
            lstMachComboBoxes[4].Text = "";
            lstMachComboBoxes[4].Visibility = Visibility.Hidden;
            lstProgComboBoxes[4].SelectedIndex = 0;
            lstProgComboBoxes[4].Visibility = Visibility.Hidden;
            lstProgDateLabels[4].Content = "";
            lstProgDateLabels[4].Visibility = Visibility.Hidden;
            lstProgTextBoxes[4].Text = "";
            lstProgTextBoxes[4].Visibility = Visibility.Hidden;
            btnRemoveProgram4.Visibility = Visibility.Hidden;
            btnAdd4.Visibility = Visibility.Visible;
            btnAdd5.Visibility = Visibility.Hidden;
            btnRemoveProgram3.Visibility = Visibility.Visible;
        }

        private void btnRemoveProgram5_Click(object sender, RoutedEventArgs e)
        {
            if (ExportList.Count > 5)
                ExportList.RemoveAt(5);
            lstMachComboBoxes[5].Text = "";
            lstMachComboBoxes[5].Visibility = Visibility.Hidden;
            lstProgComboBoxes[5].SelectedIndex = 0;
            lstProgComboBoxes[5].Visibility = Visibility.Hidden;
            lstProgDateLabels[5].Content = "";
            lstProgDateLabels[5].Visibility = Visibility.Hidden;
            lstProgTextBoxes[5].Text = "";
            lstProgTextBoxes[5].Visibility = Visibility.Hidden;
            btnRemoveProgram5.Visibility = Visibility.Hidden;
            btnAdd5.Visibility = Visibility.Visible;
            btnAdd6.Visibility = Visibility.Hidden;
            btnRemoveProgram4.Visibility = Visibility.Visible;
        }

        private void btnRemoveProgram6_Click(object sender, RoutedEventArgs e)
        {
            if (ExportList.Count > 6)
                ExportList.RemoveAt(6);
            lstMachComboBoxes[6].Text = "";
            lstMachComboBoxes[6].Visibility = Visibility.Hidden;
            lstProgComboBoxes[6].SelectedIndex = 0;
            lstProgComboBoxes[6].Visibility = Visibility.Hidden;
            lstProgDateLabels[6].Content = "";
            lstProgDateLabels[6].Visibility = Visibility.Hidden;
            lstProgTextBoxes[6].Text = "";
            lstProgTextBoxes[6].Visibility = Visibility.Hidden;
            btnRemoveProgram6.Visibility = Visibility.Hidden;
            btnAdd6.Visibility = Visibility.Visible;
            btnAdd7.Visibility = Visibility.Hidden;
            btnRemoveProgram5.Visibility = Visibility.Visible;
        }

        private void btnRemoveProgram7_Click(object sender, RoutedEventArgs e)
        {
            if (ExportList.Count > 7)
                ExportList.RemoveAt(7);
            lstMachComboBoxes[7].Text = "";
            lstMachComboBoxes[7].Visibility = Visibility.Hidden;
            lstProgComboBoxes[7].SelectedIndex = 0;
            lstProgComboBoxes[7].Visibility = Visibility.Hidden;
            lstProgDateLabels[7].Content = "";
            lstProgDateLabels[7].Visibility = Visibility.Hidden;
            lstProgTextBoxes[7].Text = "";
            lstProgTextBoxes[7].Visibility = Visibility.Hidden;
            btnRemoveProgram7.Visibility = Visibility.Hidden;
            btnAdd7.Visibility = Visibility.Visible;
            btnRemoveProgram6.Visibility = Visibility.Visible;
        }

        private void textBoxProg1_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".ci2";
            dlg.Filter = "Universal Export|*.ci2|All Files|*.*";
            dlg.Title = "Load a Program Export";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                string file = dlg.FileName;

                FileParserFactory factory = new Ci2ParserFactory(file);
                Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                if (parser != null)
                {
                    if (ExportList.Count > 0)
                        ExportList[0] = parser;
                    else
                        ExportList.Add(parser);
                    (sender as TextBox).Text = parser.ProgramName;
                    lstProgDateLabels[0].Content = parser.DateCreated;
                    AutoDetectExportInfo();
                    if (parser.Pass.Equals(lstPasses[1]))
                        lstProgComboBoxes[0].SelectedIndex = 1;
                    else
                        lstProgComboBoxes[0].SelectedIndex = 0;
                    foreach (var item in dictMachines)
                    {
                        if (parser.MachineName.Equals(item.Value))
                            lstMachComboBoxes[0].SelectedIndex = item.Key;
                    }

                }
            }
        }

        private void textBoxProg2_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ExportList.Count < 1)
                return;
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".ci2";
            dlg.Filter = "Universal Export|*.ci2|All Files|*.*";
            dlg.Title = "Load a Program Export";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                string file = dlg.FileName;

                FileParserFactory factory = new Ci2ParserFactory(file);
                Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                if (parser != null)
                {
                    if (ExportList.Count > 1)
                        ExportList[1] = parser;
                    else
                        ExportList.Add(parser);
                    (sender as TextBox).Text = parser.ProgramName;
                    lstProgDateLabels[1].Content = parser.DateCreated;
                    AutoDetectExportInfo();
                    if (parser.Pass.Equals(lstPasses[1]))
                        lstProgComboBoxes[1].SelectedIndex = 1;
                    else
                        lstProgComboBoxes[1].SelectedIndex = 0;
                    foreach (var item in dictMachines)
                    {
                        if (parser.MachineName.Equals(item.Value))
                            lstMachComboBoxes[1].SelectedIndex = item.Key;
                    }

                }
            }
        }

        private void textBoxProg3_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ExportList.Count < 2)
                return;
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".ci2";
            dlg.Filter = "Universal Export|*.ci2|All Files|*.*";
            dlg.Title = "Load a Program Export";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                string file = dlg.FileName;

                FileParserFactory factory = new Ci2ParserFactory(file);
                Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                if (parser != null)
                {
                    if (ExportList.Count > 2)
                        ExportList[2] = parser;
                    else
                        ExportList.Add(parser);                 
                    (sender as TextBox).Text = parser.ProgramName;
                    lstProgDateLabels[2].Content = parser.DateCreated;
                    AutoDetectExportInfo();
                    if (parser.Pass.Equals(lstPasses[1]))
                        lstProgComboBoxes[2].SelectedIndex = 1;
                    else
                        lstProgComboBoxes[2].SelectedIndex = 0;
                    foreach (var item in dictMachines)
                    {
                        if (parser.MachineName.Equals(item.Value))
                            lstMachComboBoxes[2].SelectedIndex = item.Key;
                    }

                }
            }
        }

        private void textBoxProg4_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ExportList.Count < 3)
                return;
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".ci2";
            dlg.Filter = "Universal Export|*.ci2|All Files|*.*";
            dlg.Title = "Load a Program Export";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                string file = dlg.FileName;

                FileParserFactory factory = new Ci2ParserFactory(file);
                Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                if (parser != null)
                {
                    if (ExportList.Count > 3)
                        ExportList[3] = parser;
                    else
                        ExportList.Add(parser);
                    (sender as TextBox).Text = parser.ProgramName;
                    lstProgDateLabels[3].Content = parser.DateCreated;
                    AutoDetectExportInfo();
                    if (parser.Pass.Equals(lstPasses[1]))
                        lstProgComboBoxes[3].SelectedIndex = 1;
                    else
                        lstProgComboBoxes[3].SelectedIndex = 0;
                    foreach (var item in dictMachines)
                    {
                        if (parser.MachineName.Equals(item.Value))
                            lstMachComboBoxes[3].SelectedIndex = item.Key;
                    }

                }
            }
        }

        private void textBoxProg5_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ExportList.Count < 4)
                return;
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".ci2";
            dlg.Filter = "Universal Export|*.ci2|All Files|*.*";
            dlg.Title = "Load a Program Export";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                string file = dlg.FileName;

                FileParserFactory factory = new Ci2ParserFactory(file);
                Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                if (parser != null)
                {
                    if (ExportList.Count > 4)
                        ExportList[4] = parser;
                    else
                        ExportList.Add(parser);
                    (sender as TextBox).Text = parser.ProgramName;
                    lstProgDateLabels[4].Content = parser.DateCreated;
                    AutoDetectExportInfo();
                    if (parser.Pass.Equals(lstPasses[1]))
                        lstProgComboBoxes[4].SelectedIndex = 1;
                    else
                        lstProgComboBoxes[4].SelectedIndex = 0;
                    foreach (var item in dictMachines)
                    {
                        if (parser.MachineName.Equals(item.Value))
                            lstMachComboBoxes[4].SelectedIndex = item.Key;
                    }

                }
            }
        }

        private void textBoxProg6_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ExportList.Count < 5)
                return;
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".ci2";
            dlg.Filter = "Universal Export|*.ci2|All Files|*.*";
            dlg.Title = "Load a Program Export";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                string file = dlg.FileName;

                FileParserFactory factory = new Ci2ParserFactory(file);
                Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                if (parser != null)
                {
                    if (ExportList.Count > 5)
                        ExportList[5] = parser;
                    else
                        ExportList.Add(parser);
                    (sender as TextBox).Text = parser.ProgramName;
                    lstProgDateLabels[5].Content = parser.DateCreated;
                    AutoDetectExportInfo();
                    if (parser.Pass.Equals(lstPasses[1]))
                        lstProgComboBoxes[5].SelectedIndex = 1;
                    else
                        lstProgComboBoxes[5].SelectedIndex = 0;
                    foreach (var item in dictMachines)
                    {
                        if (parser.MachineName.Equals(item.Value))
                            lstMachComboBoxes[5].SelectedIndex = item.Key;
                    }

                }
            }
        }

        private void textBoxProg7_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ExportList.Count < 6)
                return;
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".ci2";
            dlg.Filter = "Universal Export|*.ci2|All Files|*.*";
            dlg.Title = "Load a Program Export";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                string file = dlg.FileName;

                FileParserFactory factory = new Ci2ParserFactory(file);
                Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                if (parser != null)
                {
                    if (ExportList.Count > 6)
                        ExportList[6] = parser;
                    else
                        ExportList.Add(parser);
                    (sender as TextBox).Text = parser.ProgramName;
                    lstProgDateLabels[6].Content = parser.DateCreated;
                    AutoDetectExportInfo();
                    if (parser.Pass.Equals(lstPasses[1]))
                        lstProgComboBoxes[6].SelectedIndex = 1;
                    else
                        lstProgComboBoxes[6].SelectedIndex = 0;
                    foreach (var item in dictMachines)
                    {
                        if (parser.MachineName.Equals(item.Value))
                            lstMachComboBoxes[6].SelectedIndex = item.Key;
                    }

                }
            }
        }

        private void textBoxProg8_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (ExportList.Count < 7)
                return;
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".ci2";
            dlg.Filter = "Universal Export|*.ci2|All Files|*.*";
            dlg.Title = "Load a Program Export";
            bool? result = dlg.ShowDialog();
            if (result == true)
            {
                string file = dlg.FileName;

                FileParserFactory factory = new Ci2ParserFactory(file);
                Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                if (parser != null)
                {
                    if (ExportList.Count > 7)
                        ExportList[7] = parser;
                    else
                        ExportList.Add(parser);
                    (sender as TextBox).Text = parser.ProgramName;
                    lstProgDateLabels[7].Content = parser.DateCreated;
                    AutoDetectExportInfo();
                    if (parser.Pass.Equals(lstPasses[1]))
                        lstProgComboBoxes[7].SelectedIndex = 1;
                    else
                        lstProgComboBoxes[7].SelectedIndex = 0;
                    foreach (var item in dictMachines)
                    {
                        if (parser.MachineName.Equals(item.Value))
                            lstMachComboBoxes[7].SelectedIndex = item.Key;
                    }

                }
            }
        }

        private void textBoxBom_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            LoadBOM();
            if (Bom != null)
                btnCreate.Visibility = Visibility.Visible;
        }

        private void btnFlipLine_Click(object sender, RoutedEventArgs e)
        {
            SwapUIBoxes();
        }

        private void comparebutton_Click(object sender, RoutedEventArgs e)
        {
            List<Ci2Parser> oldCi2s = new List<Ci2Parser>();
            List<Ci2Parser> newCi2s = new List<Ci2Parser>();
            SetupSheetParser oldSsParser = null;
            SetupSheetParser newSsParser = null;

            ///Load old ci2s
            ///
            Microsoft.Win32.OpenFileDialog oldCi2dlg = new Microsoft.Win32.OpenFileDialog();
            oldCi2dlg.DefaultExt = ".ci2";
            oldCi2dlg.Filter = "Old Universal Exports|*.ci2|All Files|*.*";
            oldCi2dlg.Title = "Load the Old Program Exports";
            oldCi2dlg.Multiselect = true;
            bool? result = oldCi2dlg.ShowDialog();
            if (result == true)
            {
                string[] files = oldCi2dlg.FileNames;
                foreach (var filename in files)
                {
                    FileParserFactory factory = new Ci2ParserFactory(filename);
                    Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                    if (parser != null)
                    {
                        if (!parser.IsValid)
                        {
                            MessageBox.Show("Could not create ci2 export object. Check input files and try again.", "comparebutton_Click()",
                                MessageBoxButton.OK, MessageBoxImage.Error);
                            return;
                        }
                        oldCi2s.Add(parser);
                    }
                }
                ///Load old setup sheet
                ///
                Microsoft.Win32.OpenFileDialog oldSetupDlg = new Microsoft.Win32.OpenFileDialog();
                oldSetupDlg.DefaultExt = ".rtf";
                oldSetupDlg.Filter = "Old Setup Sheet|*.rtf|All Files|*.*";
                oldSetupDlg.Title = "Load the Old Setup Sheet";
                bool? result2 = oldSetupDlg.ShowDialog();
                if (result2 == true)
                {
                    FileParserFactory factory = new SetupSheetParserFactory(oldSetupDlg.FileName);
                    oldSsParser = (SetupSheetParser)factory.GetFileParser();
                    if (!oldSsParser.IsValid)
                    {
                        MessageBox.Show("Could not create setup sheet object. Check input files and try again.", "comparebutton_Click()",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                }
                else
                {
                    MessageBox.Show("bye!", "Bye", MessageBoxButton.OK, MessageBoxImage.None);
                    return;
                }

                //
                //Load new ci2s
                Microsoft.Win32.OpenFileDialog newCi2dlg = new Microsoft.Win32.OpenFileDialog();
                newCi2dlg.DefaultExt = ".ci2";
                newCi2dlg.Filter = "New Universal Exports|*.ci2|All Files|*.*";
                newCi2dlg.Title = "Load the New Program Exports";
                newCi2dlg.Multiselect = true;
                bool? result3 = newCi2dlg.ShowDialog();
                if (result3 == true)
                {
                    string[] files2 = newCi2dlg.FileNames;
                    foreach (var filename in files2)
                    {
                        FileParserFactory factory = new Ci2ParserFactory(filename);
                        Ci2Parser parser = (Ci2Parser)factory.GetFileParser();
                        if (parser != null)
                        {
                            if (!parser.IsValid)
                            {
                                MessageBox.Show("Could not create ci2 export object. Check input files and try again.", "comparebutton_Click()",
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                                return;
                            }
                            newCi2s.Add(parser);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("bye!", "Bye", MessageBoxButton.OK, MessageBoxImage.None);
                    return;
                }
                ///Load new setup sheet
                ///
                Microsoft.Win32.OpenFileDialog newSetupDlg = new Microsoft.Win32.OpenFileDialog();
                newSetupDlg.DefaultExt = ".rtf";
                newSetupDlg.Filter = "New Setup Sheet|*.rtf|All Files|*.*";
                newSetupDlg.Title = "Load the New Setup Sheet";
                bool? result4 = newSetupDlg.ShowDialog();
                if (result4 == true)
                {
                    FileParserFactory factory = new SetupSheetParserFactory(newSetupDlg.FileName);
                    newSsParser = (SetupSheetParser)factory.GetFileParser();
                    if (!newSsParser.IsValid)
                    {
                        MessageBox.Show("Could not create setup sheet object. Check input files and try again.", "comparebutton_Click()",
                            MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }

                }
                else
                {
                    MessageBox.Show("bye!", "Bye", MessageBoxButton.OK, MessageBoxImage.None);
                    return;
                }
            }
            else
                return;
            ///Do the Compare
            ///
            ExportAndSetupComparer Comparer = new ExportAndSetupComparer(oldSsParser, newSsParser, oldCi2s, newCi2s);
            Comparer.CompareSetupSheets();
            if(Comparer.Valid)
                Comparer.SaveResults();
        }

        private void famcheckbutton_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Not yet implemented. Coming eventually...", "Sorry :(", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
}
