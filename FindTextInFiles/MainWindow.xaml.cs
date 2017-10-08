using System.Windows;
using IdealAutomate.Core;
using System.Collections.Generic;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Collections;
using System.Windows.Controls;
using System.Configuration;
using WindowsInput;
using WindowsInput.Native;
using System.Threading;


namespace FindTextInFiles {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>

    public partial class MainWindow : Window {
        public static string strPathToSearch = @"C:\SVNIA\trunk";

        public static string strSearchPattern = @"*.*";

        public static string strSearchExcludePattern = @"*.dll;*.exe;*.png;*.xml;*.cache;*.sln;*.suo;*.pdb;*.csproj;*.deploy";

        public static string strSearchText = @"notepad";

        public static string strLowerCaseSearchText = @"notepad";

        public static int intHits;

        public static bool boolMatchCase = false;

        public static bool boolUseRegularExpression = false;

        public static bool boolStringFoundInFile;
        string strFindWhat = "";

        public static List<MatchInfo> matchInfoList;
        public static List<HotKeyRecord> listHotKeyRecords;
        public static Dictionary<string, VirtualKeyCode> dictVirtualKeyCodes = new Dictionary<string, VirtualKeyCode>();
        bool boolStopEvent = false;
        IdealAutomate.Core.Methods myActions;

        public MainWindow() {

            bool boolRunningFromHome = false;
            var window = new Window() //make sure the window is invisible
            {
                Width = 0,
                Height = 0,
                Left = -2000,
                WindowStyle = WindowStyle.None,
                ShowInTaskbar = false,
                ShowActivated = false,
            };

            window.Show();
            myActions = new Methods();
            myActions.ScriptStartedUpdateStats();

            string strSavedDomainName = myActions.GetValueByKey("DomainName");
            if (strSavedDomainName == "") {
                strSavedDomainName = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
            }
            InitializeComponent();
            this.Hide();


            listHotKeyRecords = new List<HotKeyRecord>();

            HotKeyRecord myHotKeyRecord = new HotKeyRecord();
            string strHotKey = "Ctrl+Alt+N";
            myHotKeyRecord.HotKeys = strHotKey.Split('+');
            myHotKeyRecord.Executable = @"OpenLineInNotepad";
            myHotKeyRecord.ExecuteContent = null;
            myHotKeyRecord.ScriptID = 0;
            bool boolHotKeysGood = true;
            foreach (string myHotKey in myHotKeyRecord.HotKeys) {
                if (dictVirtualKeyCodes.ContainsKey(myHotKey)) {
                    MessageBox.Show("Invalid hotkey: " + myHotKey);
                    boolHotKeysGood = false;
                }
            }
            if (boolHotKeysGood) {
                listHotKeyRecords.Add(myHotKeyRecord);
            }

            dictVirtualKeyCodes.Add("Ctrl", VirtualKeyCode.CONTROL);
            dictVirtualKeyCodes.Add("Alt", VirtualKeyCode.MENU);
            dictVirtualKeyCodes.Add("Shift", VirtualKeyCode.SHIFT);
            dictVirtualKeyCodes.Add("Space", VirtualKeyCode.SPACE);
            dictVirtualKeyCodes.Add("Up", VirtualKeyCode.UP);
            dictVirtualKeyCodes.Add("Down", VirtualKeyCode.DOWN);
            dictVirtualKeyCodes.Add("Left", VirtualKeyCode.LEFT);
            dictVirtualKeyCodes.Add("Right", VirtualKeyCode.RIGHT);
            dictVirtualKeyCodes.Add("A", VirtualKeyCode.VK_A);
            dictVirtualKeyCodes.Add("B", VirtualKeyCode.VK_B);
            dictVirtualKeyCodes.Add("C", VirtualKeyCode.VK_C);
            dictVirtualKeyCodes.Add("D", VirtualKeyCode.VK_D);
            dictVirtualKeyCodes.Add("E", VirtualKeyCode.VK_E);
            dictVirtualKeyCodes.Add("F", VirtualKeyCode.VK_F);
            dictVirtualKeyCodes.Add("G", VirtualKeyCode.VK_G);
            dictVirtualKeyCodes.Add("H", VirtualKeyCode.VK_H);
            dictVirtualKeyCodes.Add("I", VirtualKeyCode.VK_I);
            dictVirtualKeyCodes.Add("J", VirtualKeyCode.VK_J);
            dictVirtualKeyCodes.Add("K", VirtualKeyCode.VK_K);
            dictVirtualKeyCodes.Add("L", VirtualKeyCode.VK_L);
            dictVirtualKeyCodes.Add("M", VirtualKeyCode.VK_M);
            dictVirtualKeyCodes.Add("N", VirtualKeyCode.VK_N);
            dictVirtualKeyCodes.Add("O", VirtualKeyCode.VK_O);
            dictVirtualKeyCodes.Add("P", VirtualKeyCode.VK_P);
            dictVirtualKeyCodes.Add("Q", VirtualKeyCode.VK_Q);
            dictVirtualKeyCodes.Add("R", VirtualKeyCode.VK_R);
            dictVirtualKeyCodes.Add("S", VirtualKeyCode.VK_S);
            dictVirtualKeyCodes.Add("T", VirtualKeyCode.VK_T);
            dictVirtualKeyCodes.Add("U", VirtualKeyCode.VK_U);
            dictVirtualKeyCodes.Add("V", VirtualKeyCode.VK_V);
            dictVirtualKeyCodes.Add("W", VirtualKeyCode.VK_W);
            dictVirtualKeyCodes.Add("X", VirtualKeyCode.VK_X);
            dictVirtualKeyCodes.Add("Y", VirtualKeyCode.VK_Y);
            dictVirtualKeyCodes.Add("Z", VirtualKeyCode.VK_Z);
            // Create a timer and set a two millisecond interval.
            System.Timers.Timer aTimer = new System.Timers.Timer();
            aTimer.Interval = 2;

            // Alternate method: create a Timer with an interval argument to the constructor. 
            //aTimer = new System.Timers.Timer(2000); 

            // Create a timer with a two millisecond interval.
            aTimer = new System.Timers.Timer(2);

            // Hook up the Elapsed event for the timer. 
            aTimer.Elapsed += OnTimedEvent;

            // Have the timer fire repeated events (true is the default)
            aTimer.AutoReset = true;

            // Start the timer
            aTimer.Enabled = true;
            DisplayFindTextInFilesWindow:
            int intRowCtr = 0;
            ControlEntity myControlEntity = new ControlEntity();
            List<ControlEntity> myListControlEntity = new List<ControlEntity>();
            List<ComboBoxPair> cbp = new List<ComboBoxPair>();
            List<ComboBoxPair> cbp1 = new List<ComboBoxPair>();
            List<ComboBoxPair> cbp2 = new List<ComboBoxPair>();
            List<ComboBoxPair> cbp3 = new List<ComboBoxPair>();
            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.Heading;
            myControlEntity.ID = "lbl";
            myControlEntity.Text = "Find Text In Files";
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.ColumnNumber = 0;
            myListControlEntity.Add(myControlEntity.CreateControlEntity());

            intRowCtr++;
            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.Label;
            myControlEntity.ID = "lblFindWhat";
            myControlEntity.Text = "FindWhat";
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.Width = 150;
            myControlEntity.ColumnNumber = 0;
            myControlEntity.ColumnSpan = 1;
            myListControlEntity.Add(myControlEntity.CreateControlEntity());



            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.ComboBox;
            myControlEntity.SelectedValue = myActions.GetValueByKey("cbxFindWhatSelectedValue");
            myControlEntity.ID = "cbxFindWhat";
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.ToolTipx = "";
            //foreach (var item in alcbxFindWhat) {
            //    cbp.Add(new ComboBoxPair(item.ToString(), item.ToString()));
            //}
            //myControlEntity.ListOfKeyValuePairs = cbp;
            myControlEntity.ComboBoxIsEditable = true;
            myControlEntity.ColumnNumber = 1;

            myControlEntity.ColumnSpan = 2;
            myListControlEntity.Add(myControlEntity.CreateControlEntity());


            intRowCtr++;
            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.Label;
            myControlEntity.ID = "lblFileType";
            myControlEntity.Text = "FileType";
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.Width = 150;
            myControlEntity.ColumnNumber = 0;
            myControlEntity.ColumnSpan = 1;
            myListControlEntity.Add(myControlEntity.CreateControlEntity());


            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.ComboBox;
            myControlEntity.SelectedValue = myActions.GetValueByKey("cbxFileTypeSelectedValue");
            myControlEntity.ID = "cbxFileType";
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.ToolTipx = "Here is an example: *.*";
            //foreach (var item in alcbxFileType) {
            //    cbp1.Add(new ComboBoxPair(item.ToString(), item.ToString()));
            //}
            //myControlEntity.ListOfKeyValuePairs = cbp1;
            myControlEntity.ComboBoxIsEditable = true;
            myControlEntity.ColumnNumber = 1;
            myControlEntity.ColumnSpan = 2;
            myListControlEntity.Add(myControlEntity.CreateControlEntity());


            intRowCtr++;
            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.Label;
            myControlEntity.ID = "lblExclude";
            myControlEntity.Text = "Exclude";
            myControlEntity.Width = 150;
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.ColumnNumber = 0;
            myControlEntity.ColumnSpan = 1;
            myListControlEntity.Add(myControlEntity.CreateControlEntity());


            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.ComboBox;
            myControlEntity.SelectedValue = myActions.GetValueByKey("cbxExcludeSelectedValue");
            myControlEntity.ID = "cbxExclude";
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.ToolTipx = "Here is an example: *.dll;*.exe;*.png;*.xml;*.cache;*.sln;*.suo;*.pdb;*.csproj;*.deploy";
            myControlEntity.ComboBoxIsEditable = true;
            myControlEntity.ColumnNumber = 1;
            myControlEntity.ColumnSpan = 2;
            myListControlEntity.Add(myControlEntity.CreateControlEntity());


            intRowCtr++;
            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.Label;
            myControlEntity.ID = "lblFolder";
            myControlEntity.Text = "Folder";
            myControlEntity.Width = 150;
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.ColumnNumber = 0;
            myControlEntity.ColumnSpan = 1;
            myListControlEntity.Add(myControlEntity.CreateControlEntity());

            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.ComboBox;
            myControlEntity.SelectedValue = myActions.GetValueByKey("cbxFolderSelectedValue");
            myControlEntity.ID = "cbxFolder";
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.ToolTipx = @"Here is an example: C:\Users\harve\Documents\GitHub";
            myControlEntity.ComboBoxIsEditable = true;
            myControlEntity.ColumnNumber = 1;
            myControlEntity.ColumnSpan = 2;
            myListControlEntity.Add(myControlEntity.CreateControlEntity());

            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.Button;
            myControlEntity.ID = "btnSelectFolder";
            myControlEntity.Text = "Select Folder...";
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.ColumnNumber = 3;
            myListControlEntity.Add(myControlEntity.CreateControlEntity());

            intRowCtr++;
            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.CheckBox;
            myControlEntity.ID = "chkMatchCase";
            myControlEntity.Text = "Match Case";
            myControlEntity.Width = 150;
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.ColumnNumber = 0;
            myControlEntity.ColumnSpan = 1;
            string strMatchCase = myActions.GetValueByKey("chkMatchCase");

            if (strMatchCase.ToLower() == "true") {
                myControlEntity.Checked = true;
            } else {
                myControlEntity.Checked = false;
            }
            myListControlEntity.Add(myControlEntity.CreateControlEntity());

            intRowCtr++;
            myControlEntity.ControlEntitySetDefaults();
            myControlEntity.ControlType = ControlType.CheckBox;
            myControlEntity.ID = "chkUseRegularExpression";
            myControlEntity.Text = "UseRegularExpression";
            myControlEntity.Width = 150;
            myControlEntity.RowNumber = intRowCtr;
            myControlEntity.ColumnNumber = 0;
            myControlEntity.ColumnSpan = 1;
            string strUseRegularExpression = myActions.GetValueByKey("chkUseRegularExpression");
            if (strUseRegularExpression.ToLower() == "true") {
                myControlEntity.Checked = true;
            } else {
                myControlEntity.Checked = false;
            }
            myListControlEntity.Add(myControlEntity.CreateControlEntity());

            DisplayWindowAgain:
            string strButtonPressed = myActions.WindowMultipleControls(ref myListControlEntity, 300, 1200, 100, 100);
            LineAfterDisplayWindow:
            if (strButtonPressed == "btnCancel") {
                myActions.MessageBoxShow("Okay button not pressed - Script Cancelled");
                goto myExit;
            }

            boolMatchCase = myListControlEntity.Find(x => x.ID == "chkMatchCase").Checked;
            boolUseRegularExpression = myListControlEntity.Find(x => x.ID == "chkUseRegularExpression").Checked;

            strFindWhat = myListControlEntity.Find(x => x.ID == "cbxFindWhat").SelectedValue;
            //  string strFindWhatKey = myListControlEntity.Find(x => x.ID == "cbxFindWhat").SelectedKey;

            string strFileType = myListControlEntity.Find(x => x.ID == "cbxFileType").SelectedValue;
            //     string strFileTypeKey = myListControlEntity.Find(x => x.ID == "cbxFileType").SelectedKey;

            string strExclude = myListControlEntity.Find(x => x.ID == "cbxExclude").SelectedValue;
            //      string strExcludeKey = myListControlEntity.Find(x => x.ID == "cbxExclude").SelectedKey;

            string strFolder = myListControlEntity.Find(x => x.ID == "cbxFolder").SelectedValue;
            //     string strFolderKey = myListControlEntity.Find(x => x.ID == "cbxFolder").SelectedKey;
            myActions.SetValueByKey("chkMatchCase", boolMatchCase.ToString());
            myActions.SetValueByKey("chkUseRegularExpression", boolUseRegularExpression.ToString());
            myActions.SetValueByKey("cbxFindWhatSelectedValue", strFindWhat);
            myActions.SetValueByKey("cbxFileTypeSelectedValue", strFileType);
            myActions.SetValueByKey("cbxExcludeSelectedValue", strExclude);
            myActions.SetValueByKey("cbxFolderSelectedValue", strFolder);
            string settingsDirectory = "";
            if (strButtonPressed == "btnSelectFolder") {
                var dialog = new System.Windows.Forms.FolderBrowserDialog();
                dialog.SelectedPath = myActions.GetValueByKey("LastSearchFolder");
                string str = "LastSearchFolder";


                System.Windows.Forms.DialogResult result = dialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && Directory.Exists(dialog.SelectedPath)) {
                    myListControlEntity.Find(x => x.ID == "cbxFolder").SelectedValue = dialog.SelectedPath;
                    myListControlEntity.Find(x => x.ID == "cbxFolder").SelectedKey = dialog.SelectedPath;
                    myListControlEntity.Find(x => x.ID == "cbxFolder").Text = dialog.SelectedPath;
                    myActions.SetValueByKey("LastSearchFolder", dialog.SelectedPath);
                    strFolder = dialog.SelectedPath;
                    myActions.SetValueByKey("cbxFolderSelectedValue", strFolder);
                    string strScriptName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;
                    string fileName = "cbxFolder.txt";
                    string strApplicationBinDebug = System.Windows.Forms.Application.StartupPath;
                    string myNewProjectSourcePath = strApplicationBinDebug.Replace("\\bin\\Debug", "");
                    settingsDirectory = GetAppDirectoryForScript(myActions.ConvertFullFileNameToScriptPath(myNewProjectSourcePath));
                    string settingsPath = System.IO.Path.Combine(settingsDirectory, fileName);
                    ArrayList alHosts = new ArrayList();
                    cbp = new List<ComboBoxPair>();
                    cbp.Clear();
                    cbp.Add(new ComboBoxPair("--Select Item ---", "--Select Item ---"));
                    ComboBox myComboBox = new ComboBox();


                    if (!File.Exists(settingsPath)) {
                        using (StreamWriter objSWFile = File.CreateText(settingsPath)) {
                            objSWFile.Close();
                        }
                    }
                    using (StreamReader objSRFile = File.OpenText(settingsPath)) {
                        string strReadLine = "";
                        while ((strReadLine = objSRFile.ReadLine()) != null) {
                            string[] keyvalue = strReadLine.Split('^');
                            if (keyvalue[0] != "--Select Item ---") {
                                cbp.Add(new ComboBoxPair(keyvalue[0], keyvalue[1]));
                            }
                        }
                        objSRFile.Close();
                    }
                    string strNewHostName = dialog.SelectedPath;
                    List<ComboBoxPair> alHostx = cbp;
                    List<ComboBoxPair> alHostsNew = new List<ComboBoxPair>();
                    ComboBoxPair myCbp = new ComboBoxPair(strNewHostName, strNewHostName);
                    bool boolNewItem = false;

                    alHostsNew.Add(myCbp);
                    if (alHostx.Count > 14) {
                        for (int i = alHostx.Count - 1; i > 0; i--) {
                            if (alHostx[i]._Key.Trim() != "--Select Item ---") {
                                alHostx.RemoveAt(i);
                                break;
                            }
                        }
                    }
                    foreach (ComboBoxPair item in alHostx) {
                        if (strNewHostName != item._Key && item._Key != "--Select Item ---") {
                            boolNewItem = true;
                            alHostsNew.Add(item);
                        }
                    }

                    using (StreamWriter objSWFile = File.CreateText(settingsPath)) {
                        foreach (ComboBoxPair item in alHostsNew) {
                            if (item._Key != "") {
                                objSWFile.WriteLine(item._Key + '^' + item._Value);
                            }
                        }
                        objSWFile.Close();
                    }
                    goto DisplayWindowAgain;
                }
            }
            string strFindWhatToUse = "";
            string strFileTypeToUse = "";
            string strExcludeToUse = "";
            string strFolderToUse = "";
            if (strButtonPressed == "btnOkay") {
                if ((strFindWhat == "--Select Item ---" || strFindWhat == "")) {
                    myActions.MessageBoxShow("Please enter Find What or select Find What from ComboBox; else press Cancel to Exit");
                    goto DisplayFindTextInFilesWindow;
                }
                if ((strFileType == "--Select Item ---" || strFileType == "")) {
                    myActions.MessageBoxShow("Please enter File Type or select File Type from ComboBox; else press Cancel to Exit");
                    goto DisplayFindTextInFilesWindow;
                }
                if ((strExclude == "--Select Item ---" || strExclude == "")) {
                    myActions.MessageBoxShow("Please enter Exclude or select Exclude from ComboBox; else press Cancel to Exit");
                    goto DisplayFindTextInFilesWindow;
                }
                if ((strFolder == "--Select Item ---" || strFolder == "")) {
                    myActions.MessageBoxShow("Please enter Folder or select Folder from ComboBox; else press Cancel to Exit");
                    goto DisplayFindTextInFilesWindow;
                }



                strFindWhatToUse = strFindWhat;

                if (boolUseRegularExpression) {
                    strFindWhatToUse = strFindWhatToUse.Replace(")", "\\)").Replace("(", "\\(");
                }


                strFileTypeToUse = strFileType;



                strExcludeToUse = strExclude;


                strFolderToUse = strFolder;


            }


            strPathToSearch = strFolderToUse;

            strSearchPattern = strFileTypeToUse;

            strSearchExcludePattern = strExcludeToUse;

            strSearchText = strFindWhatToUse;

            strLowerCaseSearchText = strFindWhatToUse.ToLower();
            myActions.SetValueByKey("FindWhatToUse", strFindWhatToUse);

            System.Diagnostics.Stopwatch st = new System.Diagnostics.Stopwatch();
            st.Start();
            intHits = 0;
            int intLineCtr;
            List<FileInfo> myFileList = TraverseTree(strSearchPattern, strPathToSearch);
            int intFiles = 0;
            matchInfoList = new List<MatchInfo>();
            //         myFileList = myFileList.OrderBy(fi => fi.FullName).ToList();
            Parallel.ForEach(myFileList, myFileInfo => {
                intLineCtr = 0;
                boolStringFoundInFile = false;
                ReadFileToString(myFileInfo.FullName, intLineCtr, matchInfoList);
                if (boolStringFoundInFile) {
                    intFiles++;
                }
            });
            matchInfoList = matchInfoList.Where(mi => mi != null).OrderBy(mi => mi.FullName).ThenBy(mi => mi.LineNumber).ToList();
            List<string> lines = new List<string>();
            foreach (var item in matchInfoList) {
                lines.Add("\"" + item.FullName + "\"(" + item.LineNumber + "," + item.LinePosition + "): " + item.LineText.Length.ToString() + " " + item.LineText.Substring(0, item.LineText.Length > 5000 ? 5000 : item.LineText.Length));

            }
           
          
            string strApplicationBinDebug1 = System.Windows.Forms.Application.StartupPath;
            string myNewProjectSourcePath1 = strApplicationBinDebug1.Replace("\\bin\\Debug", "");

            settingsDirectory = GetAppDirectoryForScript(myActions.ConvertFullFileNameToScriptPath(myNewProjectSourcePath1));
            using (FileStream fs = new FileStream(settingsDirectory + @"\MatchInfo.txt", FileMode.Create)) {
                StreamWriter file = new System.IO.StreamWriter(fs, Encoding.Default);

                file.WriteLine(@"-- " + strSearchText + " in " + strPathToSearch + " from " + strSearchPattern + " excl  " + strSearchExcludePattern + " --");
                foreach (var item in matchInfoList) {
                    file.WriteLine("\"" + item.FullName + "\"(" + item.LineNumber + "," + item.LinePosition + "): " + item.LineText.Substring(0, item.LineText.Length > 5000 ? 5000 : item.LineText.Length));
                }
                int intUniqueFiles = matchInfoList.Select(x => x.FullName).Distinct().Count();
                st.Stop();
                // Get the elapsed time as a TimeSpan value.
                TimeSpan ts = st.Elapsed;
                // Format and display the TimeSpan value.
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                    ts.Hours, ts.Minutes, ts.Seconds,
                    ts.Milliseconds / 10);
                file.WriteLine("RunTime " + elapsedTime);
                file.WriteLine(intHits.ToString() + " hits");
                // file.WriteLine(myFileList.Count().ToString() + " files");           
                file.WriteLine(intUniqueFiles.ToString() + " files with hits");
                file.Close();


                // Get the elapsed time as a TimeSpan value.
                ts = st.Elapsed;
                // Format and display the TimeSpan value.
                elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                   ts.Hours, ts.Minutes, ts.Seconds,
                   ts.Milliseconds / 10);
                Console.WriteLine("RunTime " + elapsedTime);
                Console.WriteLine(intHits.ToString() + " hits");
                // Console.WriteLine(myFileList.Count().ToString() + " files");
                Console.WriteLine(intUniqueFiles.ToString() + " files with hits");
                Console.ReadLine();
                string strExecutable = @"C:\Program Files (x86)\Notepad++\notepad++.exe";
                string strContent = settingsDirectory + @"\MatchInfo.txt";
                Process.Start(strExecutable, string.Concat("", strContent, ""));
                myActions.ScriptEndedSuccessfullyUpdateStats();
                myActions.MessageBoxShow("RunTime: " + elapsedTime + "\n\r\n\rHits: " + intHits.ToString() + "\n\r\n\rFiles with hits: " + intUniqueFiles.ToString() + "\n\r\n\rPut Cursor on line and\n\r press Ctrl+Alt+N\n\rto view detail page. ");
            }


            if (strButtonPressed == "btnOkay") {
                strButtonPressed = myActions.WindowMultipleControlsMinimized(ref myListControlEntity, 300, 1200, 100, 100);
                myActions.ScriptStartedUpdateStats();
                goto LineAfterDisplayWindow;
            }


            myExit:
            myActions.ScriptEndedSuccessfullyUpdateStats();
            Application.Current.Shutdown();
        }
        public string GetAppDirectoryForScript(string strScriptName) {
            string settingsDirectory =
      Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\IdealAutomate\\" + strScriptName;
            if (!Directory.Exists(settingsDirectory)) {
                Directory.CreateDirectory(settingsDirectory);
            }
            return settingsDirectory;
        }
        public static List<FileInfo> TraverseTree(string filterPattern, string root) {
            string[] arrayExclusionPatterns = strSearchExcludePattern.Split(';');
            for (int i = 0; i < arrayExclusionPatterns.Length; i++) {
                arrayExclusionPatterns[i] = arrayExclusionPatterns[i].ToLower().ToString().Replace("*", "");
            }
            List<FileInfo> myFileList = new List<FileInfo>();
            // Data structure to hold names of subfolders to be
            // examined for files.
            Stack<string> dirs = new Stack<string>(20);

            if (!System.IO.Directory.Exists(root)) {
                MessageBox.Show(root + " - folder did not exist");
            }


            dirs.Push(root);

            while (dirs.Count > 0) {
                string currentDir = dirs.Pop();
                string[] subDirs;
                try {
                    subDirs = System.IO.Directory.GetDirectories(currentDir);
                }
                // An UnauthorizedAccessException exception will be thrown if we do not have
                // discovery permission on a folder or file. It may or may not be acceptable 
                // to ignore the exception and continue enumerating the remaining files and 
                // folders. It is also possible (but unlikely) that a DirectoryNotFound exception 
                // will be raised. This will happen if currentDir has been deleted by
                // another application or thread after our call to Directory.Exists. The 
                // choice of which exceptions to catch depends entirely on the specific task 
                // you are intending to perform and also on how much you know with certainty 
                // about the systems on which this code will run.
                catch (UnauthorizedAccessException e) {
                    Console.WriteLine(e.Message);
                    continue;
                } catch (System.IO.DirectoryNotFoundException e) {
                    Console.WriteLine(e.Message);
                    continue;
                }

                string[] files = null;
                try {
                    files = System.IO.Directory.GetFiles(currentDir, filterPattern);
                } catch (UnauthorizedAccessException e) {

                    Console.WriteLine(e.Message);
                    continue;
                } catch (System.IO.DirectoryNotFoundException e) {
                    Console.WriteLine(e.Message);
                    continue;
                }

                // Perform the required action on each file here.
                // Modify this block to perform your required task.
                foreach (string file in files) {
                    try {
                        // Perform whatever action is required in your scenario.
                        System.IO.FileInfo fi = new System.IO.FileInfo(file);
                        bool boolFileHasGoodExtension = true;
                        foreach (var item in arrayExclusionPatterns) {
                            if (fi.FullName.ToLower().Contains(item)) {
                                boolFileHasGoodExtension = false;
                            }
                        }
                        if (boolFileHasGoodExtension) {
                            myFileList.Add(fi);
                        }
                        //    Console.WriteLine("{0}: {1}, {2}", fi.Name, fi.Length, fi.CreationTime);
                    } catch (System.IO.FileNotFoundException e) {
                        // If file was deleted by a separate application
                        //  or thread since the call to TraverseTree()
                        // then just continue.
                        Console.WriteLine(e.Message);
                        continue;
                    }
                }

                // Push the subdirectories onto the stack for traversal.
                // This could also be done before handing the files.
                foreach (string str in subDirs)
                    dirs.Push(str);
            }
            return myFileList;
        }
        public static void ReadFileToString(string fullFilePath, int intLineCtr, List<MatchInfo> matchInfoList) {
            while (true) {
                try {
                    using (FileStream fs = new FileStream(fullFilePath, FileMode.Open)) {
                        using (StreamReader sr = new StreamReader(fs, Encoding.Default)) {
                            string s;
                            string s_lower = "";
                            while ((s = sr.ReadLine()) != null) {
                                intLineCtr++;
                                if (boolUseRegularExpression) {
                                    if (boolMatchCase) {
                                        if (System.Text.RegularExpressions.Regex.IsMatch(s, strSearchText, System.Text.RegularExpressions.RegexOptions.None)) {
                                            intHits++;
                                            boolStringFoundInFile = true;
                                            MatchInfo myMatchInfo = new MatchInfo();
                                            myMatchInfo.FullName = fullFilePath;
                                            myMatchInfo.LineNumber = intLineCtr;
                                            myMatchInfo.LinePosition = s.IndexOf(strSearchText) + 1;
                                            myMatchInfo.LineText = s;
                                            matchInfoList.Add(myMatchInfo);
                                        }
                                    } else {
                                        s_lower = s.ToLower();
                                        if (System.Text.RegularExpressions.Regex.IsMatch(s_lower, strLowerCaseSearchText, System.Text.RegularExpressions.RegexOptions.IgnoreCase)) {
                                            intHits++;
                                            boolStringFoundInFile = true;
                                            MatchInfo myMatchInfo = new MatchInfo();
                                            myMatchInfo.FullName = fullFilePath;
                                            myMatchInfo.LineNumber = intLineCtr;
                                            myMatchInfo.LinePosition = s_lower.IndexOf(strLowerCaseSearchText) + 1;
                                            myMatchInfo.LineText = s;
                                            matchInfoList.Add(myMatchInfo);
                                        }
                                    }
                                } else {
                                    if (boolMatchCase) {
                                        if (s.Contains(strSearchText)) {
                                            intHits++;
                                            boolStringFoundInFile = true;
                                            MatchInfo myMatchInfo = new MatchInfo();
                                            myMatchInfo.FullName = fullFilePath;
                                            myMatchInfo.LineNumber = intLineCtr;
                                            myMatchInfo.LinePosition = s.IndexOf(strSearchText) + 1;
                                            myMatchInfo.LineText = s;
                                            matchInfoList.Add(myMatchInfo);
                                        }
                                    } else {
                                        s_lower = s.ToLower();
                                        if (s_lower.Contains(strLowerCaseSearchText)) {

                                            intHits++;
                                            boolStringFoundInFile = true;
                                            MatchInfo myMatchInfo = new MatchInfo();
                                            myMatchInfo.FullName = fullFilePath;
                                            myMatchInfo.LineNumber = intLineCtr;
                                            myMatchInfo.LinePosition = s_lower.IndexOf(strLowerCaseSearchText) + 1;
                                            myMatchInfo.LineText = s;
                                            matchInfoList.Add(myMatchInfo);
                                        }
                                    }
                                }
                            }
                            return;
                        }

                    }
                } catch (FileNotFoundException ex) {
                    Console.WriteLine("Output file {0} not yet ready ({1})", fullFilePath, ex.Message);
                    break;
                } catch (IOException ex) {
                    Console.WriteLine("Output file {0} not yet ready ({1})", fullFilePath, ex.Message);
                    break;
                } catch (UnauthorizedAccessException ex) {
                    Console.WriteLine("Output file {0} not yet ready ({1})", fullFilePath, ex.Message);
                    break;
                }
            }
        }
        private bool DoesSettingExist(string settingName) {
            return FindTextInFiles.Properties.Settings.Default.Properties.Cast<SettingsProperty>().Any(prop => prop.Name == settingName);
        }

        public void OnTimedEvent(object sender, System.Timers.ElapsedEventArgs e) {
            InputSimulator myInputSimulator = new InputSimulator();

            if (myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.CONTROL) || myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.MENU)) {
                foreach (HotKeyRecord myHotKeyRecord in listHotKeyRecords) {
                    bool boolAllHotKeysPressed = true;
                    foreach (string myHotKey in myHotKeyRecord.HotKeys) {
                        VirtualKeyCode myVirtualKeyCode;
                        dictVirtualKeyCodes.TryGetValue(myHotKey, out myVirtualKeyCode);
                        if (!myInputSimulator.InputDeviceState.IsKeyDown(myVirtualKeyCode)) {
                            boolAllHotKeysPressed = false;
                        }
                    }


                    if (boolAllHotKeysPressed && boolStopEvent == false) {
                        boolStopEvent = true;


                        if (myHotKeyRecord.Executable == "OpenLineInNotepad") {
                            OpenLineInNotepad();
                        }


                    }

                    //switch (item.HotKey.ToUpper()) {

                    //  case "P":
                    //    if (myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.VK_P)  && boolStopEvent == false) {
                    //      boolStopEvent = true;
                    //      RunWaitTillStart(item.Executable, item.ExecuteContent ?? "");

                    //      while (myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.VK_P)) {
                    //        System.Threading.Thread.Sleep(1000);
                    //      }
                    //    }
                    //    break;

                    //  case "R":
                    //    if (myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.VK_R)) {                
                    //      Run(item.Executable, item.ExecuteContent ?? "");
                    //      while (myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.CONTROL)) {
                    //        System.Threading.Thread.Sleep(200);
                    //      }
                    //    }
                    //    break;
                    //  case "S":
                    //    if (myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.VK_S)) {
                    //      Run(item.Executable, item.ExecuteContent ?? "");
                    //      while (myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.CONTROL)) {
                    //        System.Threading.Thread.Sleep(200);
                    //      }
                    //   }
                    //    break;
                    //  default:
                    //    break;
                    //}
                }
            }





            //if (myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.CONTROL) && myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.MENU) && myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.VK_S)) {


            //  Run(myActions.GetValueByKey("SVNPath","IdealAutomateDB") + ClipboardSaveToDB\ClipboardSaveToDB\bin\Debug\ClipboardSaveToDB.exe", "");
            //  while (myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.CONTROL) && myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.MENU) && myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.VK_S)) {
            //    System.Threading.Thread.Sleep(200);
            //  }


            //  //Here is the code that runs when the hotkey is pressed'
            //}
            //if (myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.CONTROL) && myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.MENU) && myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.VK_R)) {

            //  //Here is the code that runs when the hotkey is pressed'


            //  Run(myActions.GetValueByKey("SVNPath","IdealAutomateDB") + ClipboardRestoreFromDB\ClipboardRestoreFromDB\bin\Debug\ClipboardRestoreFromDB.exe", "");
            //  while (myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.CONTROL) && myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.MENU) && myInputSimulator.InputDeviceState.IsKeyDown(VirtualKeyCode.VK_R)) {
            //    System.Threading.Thread.Sleep(200);
            //  }



            //}
        }
        static void RunAsSTAThread(Action goForIt) {
            AutoResetEvent @event = new AutoResetEvent(false);
            Thread thread = new Thread(
                () => {
                    goForIt();
                    @event.Set();
                });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            @event.WaitOne();
        }
        private void OpenLineInNotepad() {
            myActions.TypeText("{RIGHT}", 500);
            myActions.TypeText("{HOME}", 500);
            myActions.TypeText("+({END})", 500);
            myActions.TypeText("^(c)", 500);
            myActions.Sleep(500);
            string strCurrentLine = "";
            RunAsSTAThread(
            () => {
                strCurrentLine = myActions.PutClipboardInEntity();
            });
            List<string> myBeginDelim = new List<string>();
            List<string> myEndDelim = new List<string>();
            myBeginDelim.Add("\"");
            myEndDelim.Add("\"");
            FindDelimitedTextParms delimParms = new FindDelimitedTextParms(myBeginDelim, myEndDelim);

            string myQuote = "\"";
            delimParms.lines[0] = strCurrentLine;


            myActions.FindDelimitedText(delimParms);
            int intLastSlash = delimParms.strDelimitedTextFound.LastIndexOf('\\');
            if (intLastSlash < 1) {
                myActions.MessageBoxShow("Could not find last slash in in EditPlusLine - aborting");
                return;
            }
            string strPathOnly = delimParms.strDelimitedTextFound.SubstringBetweenIndexes(0, intLastSlash);
            string strFileNameOnly = delimParms.strDelimitedTextFound.Substring(intLastSlash + 1);
            string strFullFileName = delimParms.strDelimitedTextFound;
            myBeginDelim.Clear();
            myEndDelim.Clear();
            myBeginDelim.Add("(");
            myEndDelim.Add(",");
            delimParms = new FindDelimitedTextParms(myBeginDelim, myEndDelim);
            delimParms.lines[0] = strCurrentLine;
            myActions.FindDelimitedText(delimParms);
            string strLineNumber = delimParms.strDelimitedTextFound;
            string strExecutable = @"C:\Program Files (x86)\Notepad++\notepad++.exe";
            string strContent = strFullFileName;
            Process.Start(strExecutable, string.Concat("", strContent, ""));
            myActions.TypeText("^(g)", 2000);
            myActions.TypeText(strLineNumber, 500);
            myActions.TypeText("{ENTER}", 500);
            myActions.TypeText("^(f)", 500);

            string strFindWhatToUse = strFindWhat;
            string blockText = strFindWhatToUse;
            strFindWhatToUse = "";
            char[] specialChars = { '{', '}', '(', ')', '+', '^' };

            foreach (char letter in blockText) {
                bool _specialCharFound = false;

                for (int i = 0; i < specialChars.Length; i++) {
                    if (letter == specialChars[i]) {
                        _specialCharFound = true;
                        break;
                    }
                }

                if (_specialCharFound)
                    strFindWhatToUse += "{" + letter.ToString() + "}";
                else
                    strFindWhatToUse += letter.ToString();
            }
            myActions.TypeText(strFindWhatToUse, 500);
            myActions.TypeText("{ENTER}", 500);
            myActions.TypeText("{ESC}", 500);
            boolStopEvent = false;
        }
    }
}