﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.IO;
using System.Linq;
using WindowsInput;
using WindowsInput.Native;



namespace IdealAutomate.Core {
    public class Methods {


        private bool fbDebugMode = true;
        private int intFileCtr = 0;
        bool boolUseGrayScaleDB = false;
        public bool DebugMode {
            get { return fbDebugMode; }
            set { fbDebugMode = value; }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern IntPtr GetOpenClipboardWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern int GetWindowText(int hwnd, StringBuilder text, int count);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int GetWindowTextLength(int hwnd);

        //Import the FindWindow API to find our window
        [DllImportAttribute("User32.dll")]
        private static extern int FindWindow(String ClassName, String WindowName);

        //Import the SetForeground API to activate it
        [DllImportAttribute("User32.dll")]
        private static extern IntPtr SetForegroundWindow(int hWnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern
          bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        private static extern
          bool IsIconic(IntPtr hWnd);

        private const int SW_HIDE = 0;
        private const int SW_SHOWNORMAL = 1;
        private const int SW_SHOWMINIMIZED = 2;
        private const int SW_SHOWMAXIMIZED = 3;
        private const int SW_SHOWNOACTIVATE = 4;
        private const int SW_RESTORE = 9;
        private const int SW_SHOWDEFAULT = 10;
        Process oProcess;
        public Methods() {

            oProcess = Process.GetCurrentProcess();

            string directory = AppDomain.CurrentDomain.BaseDirectory;
            DirectoryInfo dir = new DirectoryInfo(directory);

            foreach (FileInfo fi in dir.GetFiles()) {
                if (fi.Extension.ToUpper() == ".BMP" && fi.Name.StartsWith("temp")) {
                    fi.Delete();
                }
            }
        }




        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
        public static extern bool SetForegroundWindow(IntPtr hwnd);
        [DllImport("user32.dll")]
        static extern bool ForceForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")]
        static extern bool AllowSetForegroundWindow(int dwProcessId);
        [System.Runtime.InteropServices.DllImport("user32.dll")]

        public static extern void SwitchToThisWindow(IntPtr hWnd, bool
 fAltTab);
        [DllImport("user32")]
        static extern int BringWindowToTop(IntPtr hwnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsWindow(IntPtr hWnd);
        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);
        [DllImport("User32.DLL")]
        private static extern int AttachThreadInput(int CurrentForegroundThread, int MakeThisThreadForegrouond, bool boolAttach);




        /// <summary>
        /// <para>GetValueByKey returns the value for the specified key from KeyValueTable.</para>
        /// <para>The KeyValueTable allows you to store personal settings and</para>
        /// <para>information that you want to keep private (like passwords) in a location</para>
        /// <para>outside of your script on SQLExpress</para>
        /// </summary>
        /// <param name="pKey">The Key for the KeyValuePair</param>
        /// <param name="pInitialCatalog">usually IdealAutomateDB</param>
        /// <returns>String that is the Value for the KeyValuePair</returns>

        /// <summary>
        /// <para>WindowMultipleControls takes a list of ControlEntity objects</para>
        /// <para>and positions them in a window. When the user presses the </para>
        /// <para>okay button on the screen, the list of ControlEntity objects</para>
        /// <para>are updated with the values the user entered.  This provides</para>
        /// <para>an easy way to receive multiple values from the user</para>
        /// <para>A string is returned with the name of the button that was pressed</para>
        /// <para>Here is an example of setting background color for a button:</para>
        /// <para>myControlEntity.BackgroundColor = System.Windows.Media.Color.FromRgb(System.Drawing.Color.Red.R, System.Drawing.Color.Red.G, System.Drawing.Color.Red.B);</para>
        /// </summary>
        /// <param name="myListControlEntity">list of ControlEntity objects</param>
        /// <param name="intWindowHeight">integer indicating height of window</param>
        /// <param name="intWindowWidth">integer indicating width of window</param>
        /// <param name="intWindowTop">integer indicating number of pixels from top of screen to display window</param>
        /// <param name="intWindowLeft">integer indicating number of pixels from left side of screen to display window</param>
        /// <returns>System.Windows.Forms.DialogResult to indicate if okay button was pressed</returns>
        public string WindowMultipleControls(ref List<ControlEntity> myListControlEntity, int intWindowHeight, int intWindowWidth, int intWindowTop, int intWindowLeft) {
            if (fbDebugMode) {
                Console.WriteLine(oProcess.ProcessName + "==> " + "WindowMultipleControls");

            }
            WindowMultipleControls dlg = new WindowMultipleControls(ref myListControlEntity, intWindowHeight, intWindowWidth, intWindowTop, intWindowLeft, WindowState.Normal);

            // dlg.Owner = (Window)Window.GetWindow(this);
            // Shadow.Visibility = Visibility.Visible;
            dlg.ShowDialog();

            return dlg.strButtonClickedName;



        }
        /// <summary>
        /// <para>WindowMultipleControls takes a list of ControlEntity objects</para>
        /// <para>and positions them in a window. When the user presses the </para>
        /// <para>okay button on the screen, the list of ControlEntity objects</para>
        /// <para>are updated with the values the user entered.  This provides</para>
        /// <para>an easy way to receive multiple values from the user</para>
        /// <para>A string is returned with the name of the button that was pressed</para>
        /// <para>Here is an example of setting background color for a button:</para>
        /// <para>myControlEntity.BackgroundColor = System.Windows.Media.Color.FromRgb(System.Drawing.Color.Red.R, System.Drawing.Color.Red.G, System.Drawing.Color.Red.B);</para>
        /// </summary>
        /// <param name="myListControlEntity">list of ControlEntity objects</param>
        /// <param name="intWindowHeight">integer indicating height of window</param>
        /// <param name="intWindowWidth">integer indicating width of window</param>
        /// <param name="intWindowTop">integer indicating number of pixels from top of screen to display window</param>
        /// <param name="intWindowLeft">integer indicating number of pixels from left side of screen to display window</param>
        /// <returns>System.Windows.Forms.DialogResult to indicate if okay button was pressed</returns>
        public string WindowMultipleControlsMinimized(ref List<ControlEntity> myListControlEntity, int intWindowHeight, int intWindowWidth, int intWindowTop, int intWindowLeft) {
            if (fbDebugMode) {
                Console.WriteLine(oProcess.ProcessName + "==> " + "WindowMultipleControls");

            }
            WindowMultipleControls dlg = new WindowMultipleControls(ref myListControlEntity, intWindowHeight, intWindowWidth, intWindowTop, intWindowLeft, WindowState.Minimized);

            // dlg.Owner = (Window)Window.GetWindow(this);
            // Shadow.Visibility = Visibility.Visible;
            dlg.ShowDialog();


            return dlg.strButtonClickedName;



        }
        /// <summary>
        /// <para>WindowComboBox receives an IEnumerable of objects (ComboBoxPair) </para>
        /// <para>and a string for the label for the combobox. It returns the</para>
        /// <para>selected ComboBoxPair</para>
        /// </summary>
        /// <param name="myEntity">IEnumerable of objects</param>
        /// <param name="myEntity2">String for the label for the combobox</param>
        /// <returns>Selected ComboBoxPair</returns>


        /// <summary>
        /// <para>WindowShape allows you to display info to the user and to position that </para>
        /// <para>the window on the screen</para>
        /// </summary>
        /// <param name="myShape">string "Box" or "Arrow"</param>
        /// <param name="myOrientation">string "Left","Right","Up","Down",""</param>
        /// <param name="myTitle">string title for window</param>
        /// <param name="myContent">string content for window</param>
        /// <param name="intTop">integer indicating number of pixels from top of screen to display window</param>
        /// <param name="intLeft">integer indicating number of pixels from left of screen to display window</param>
        public void WindowShape(string myShape, string myOrientation, string myTitle, string myContent, int intTop, int intLeft) {
            if (fbDebugMode) {
                Console.WriteLine(oProcess.ProcessName + "==> " + "WindowShape: myEntity=" + myTitle);

            }
            WindowShape dlg = new WindowShape(myShape, myOrientation, myTitle, myContent, intTop, intLeft);

            // dlg.Owner = (Window)Window.GetWindow(this);
            // Shadow.Visibility = Visibility.Visible;
            dlg.ShowDialog();
            return;
        }
        /// <summary>
        /// MessageBoxShow receives an input string and displays it in a messagebox
        /// </summary>
        /// <param name="myEntity">string that you want to display in messagebox</param>
        public void MessageBoxShow(string myEntity) {
            if (fbDebugMode) {
                Console.WriteLine(oProcess.ProcessName + "==> " + "MessageBoxShow: myEntity=" + myEntity);

            }
            System.Windows.Forms.MessageBox.Show(myEntity, "Header", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.None,
          System.Windows.Forms.MessageBoxDefaultButton.Button1, (System.Windows.Forms.MessageBoxOptions)0x40000);  // MB_TOPMOST
        }
        /// <summary>
        /// MessageBoxShowWithYesNo receives an input string and displays it in a messagebox with Yes and No Buttons
        /// </summary>
        /// <param name="myEntity">string that you want to display in messagebox</param>

        /// <returns>System.Windows.Forms.DialogResult</returns>

        public void Run(string myEntityForExecutable, string myEntityForContent) {
            if (fbDebugMode) {
                Console.WriteLine(oProcess.ProcessName + "==> " + "Run: myEntityForExecutable=" + myEntityForExecutable + " myEntityForContent=" + myEntityForContent);

            }

            if (myEntityForExecutable == null) {
                string message = "Error - You need to specify executable primitive  "; // +"; EntityName is: " + myEntityForExecutable.EntityName;

                MessageBoxResult result = MessageBox.Show(message, "Run-time Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            string strExecutable = myEntityForExecutable;

            string strContent = "";
            if (myEntityForContent != null) {
                strContent = myEntityForContent;
            }
            if (strContent == "") {
                Process.Start(strExecutable);
            } else {
                try {
                    Process.Start(strExecutable, string.Concat("", strContent, ""));
                } catch (Exception ex) {

                    MessageBox.Show(ex.ToString());
                }
            }
        }
        /// <summary>
        /// <para>RunSync receives two input strings. The first is the path to the executable.</para>
        /// <para>The second is optional and it is the content you want to open with the executable.</para>
        /// <para>Run starts the executable as a thread and continues to the next statement</para>
        /// <para>AFTER the thread completes</para>
        /// </summary>
        /// <param name="myEntityForExecutable">string for the path of the executable</param>
        /// <param name="myEntityForContent">string for the content for the executable to open</param>

        public void RunSync(string myEntityForExecutable, string myEntityForContent) {
            if (fbDebugMode) {
                Console.WriteLine(oProcess.ProcessName + "==> " + "RunSync: myEntityForExecutable=" + myEntityForExecutable + " myEntityForContent=" + myEntityForContent);

            }

            if (myEntityForExecutable == null) {
                string message = "Error - You need to specify executable primitive  "; // +"; EntityName is: " + myEntityForExecutable.EntityName;
                MessageBoxResult result = MessageBox.Show(message, "Run-time Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            string strExecutable = myEntityForExecutable;

            string strContent = "";
            if (myEntityForContent != null) {
                strContent = myEntityForContent;
            }
            var p = new Process();
            p.StartInfo.FileName = strExecutable;
            if (strContent != "") {
                p.StartInfo.Arguments = string.Concat("", strContent, "");
            }
            bool started = true;
            p.Start();
            int procId = 0;
            try {
                procId = p.Id;
                Console.WriteLine("ID: " + procId);

            } catch (InvalidOperationException) {
                started = false;
            } catch (Exception ex) {
                started = false;
            }
            while (started == true && GetProcByID(procId) != null) {
                System.Threading.Thread.Sleep(1000);
            }

        }
        private Process GetProcByID(int id) {
            Process[] processlist = Process.GetProcesses();
            return processlist.FirstOrDefault(pr => pr.Id == id);
        }
        public ArrayList ReadAppDirectoryFileToArrayList(string settingsDirectory, string fileName) {
            ArrayList myArrayList = new ArrayList();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            StreamReader reader = File.OpenText(settingsPath);
            while (!reader.EndOfStream) {
                string myLine = reader.ReadLine();
                myArrayList.Add(myLine);
            }
            reader.Close();
            return myArrayList;
        }

        public string ReadValueFromAppDataFile(string settingsDirectory, string fileName) {
            StreamReader file = null;
            string strValueRead = "";
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            return strValueRead;
        }


        public void WriteValueToAppDirectoryFile(string settingsDirectory, string fileName, string strValueToWrite) {
            StreamWriter writer = null;
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            // Hook a write to the text file.
            writer = new StreamWriter(settingsPath);
            // Rewrite the entire value of s to the file
            writer.Write(strValueToWrite);
            writer.Close();
        }

        public void WriteArrayListToAppDirectoryFile(string settingsDirectory, string fileName, ArrayList arrayListToWrite) {
            StreamWriter writer = null;
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            // Hook a write to the text file.
            writer = new StreamWriter(settingsPath);
            // Rewrite the entire value of s to the file
            foreach (var item in arrayListToWrite) {
                writer.WriteLine(item.ToString());
            }
            writer.Close();
        }

        public string GetAppDirectoryForScript(string strScriptName) {
            string settingsDirectory =
      Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\IdealAutomate\\" + strScriptName;
            if (!Directory.Exists(settingsDirectory)) {
                Directory.CreateDirectory(settingsDirectory);
            }
            return settingsDirectory;
        }

        /// <summary>
        /// <para>TypeText - The visual basic SendKeys function is used to mimic </para>
        /// <para>pressing special keys (like the enter or alt keys). This means you </para>
        /// <para>need to use the Shortcut Keys help file in the IdealAutomate application </para>
        /// <para>or google in order to learn what characters can be used to represent</para>
        /// <para>special keys. For example, the ^ character is used to represent the </para>
        /// <para>control key and here is how you indicate the enter key is pressed: </para>
        /// <para>{ENTER}. You will also need to learn how to "escape" special characters</para>
        /// <para>(like the bracket character). If you are trying to type a lot of special</para>
        /// <para>characters, it may be easier to create a string primitive with the </para>
        /// <para>text you want to type and use the PutEntityInClipboard verb to copy </para>
        /// <para>the string into the clipboard. After the string is in the clipboard,</para>
        /// <para>you can use the TypeText verb with control v to paste what is in the</para>
        /// clipboard to where you want it.
        /// </summary>
        /// <param name="myEntity">string representing the keys you want to press</param>
        /// <param name="intSleep">integer representing the number of milliseconds to wait before sending the text</param>
        public void TypeText(string myEntity, int intSleep) {
            InputSimulator InputSimulator = new InputSimulator();
            //if (myEntity == "{LWin}") {
            //  KeyboardSend.KeyDown(System.Windows.Forms.Keys.ControlKey);
            //  KeyboardSend.KeyDown(System.Windows.Forms.Keys.Alt);
            //  KeyboardSend.KeyDown(System.Windows.Forms.Keys.L);
            //  KeyboardSend.KeyUp(System.Windows.Forms.Keys.L);
            //  KeyboardSend.KeyUp(System.Windows.Forms.Keys.Alt);
            //  KeyboardSend.KeyUp(System.Windows.Forms.Keys.ControlKey);
            //}
            if (intSleep > 0) {
                System.Threading.Thread.Sleep(intSleep);
            }
            //          public void SimulateSomeModifiedKeystrokes()
            //{
            //  // CTRL-C (effectively a copy command in many situations)
            //  InputSimulator.SimulateModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_C);

            //  // You can simulate chords with multiple modifiers
            //  // For example CTRL-K-C whic is simulated as
            //  // CTRL-down, K, C, CTRL-up
            //  InputSimulator.SimulateModifiedKeyStroke(VirtualKeyCode.CONTROL, new [] {VirtualKeyCode.VK_K, VirtualKeyCode.VK_C});

            //  // You can simulate complex chords with multiple modifiers and key presses
            //  // For example CTRL-ALT-SHIFT-ESC-K which is simulated as
            //  // CTRL-down, ALT-down, SHIFT-down, press ESC, press K, SHIFT-up, ALT-up, CTRL-up
            //  InputSimulator.SimulateModifiedKeyStroke(
            //    new[] { VirtualKeyCode.CONTROL, VirtualKeyCode.MENU, VirtualKeyCode.SHIFT },
            //    new[] { VirtualKeyCode.ESCAPE, VirtualKeyCode.VK_K });
            //}
            if (myEntity == "%(f)e") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_F); //System.Windows.Forms.Keys.Alt);
                InputSimulator.Keyboard.KeyPress(VirtualKeyCode.VK_E);
                return;
            }
            if (myEntity == "%(d)") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_D); //System.Windows.Forms.Keys.Alt);        
                return;
            }
            if (myEntity == "%(f)") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.VK_F); //System.Windows.Forms.Keys.Alt);        
                return;
            }
            if (myEntity == "%(\" \")n") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.SPACE); //System.Windows.Forms.Keys.Alt);
                InputSimulator.Keyboard.KeyPress(VirtualKeyCode.VK_N);
                return;
            }
            if (myEntity == "%(\" \")") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.SPACE); //System.Windows.Forms.Keys.Alt);              
                return;
            }
            if (myEntity == "%(\" \")x") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.SPACE); //System.Windows.Forms.Keys.Alt);
                InputSimulator.Keyboard.KeyPress(VirtualKeyCode.VK_X);
                return;
            }
            if (myEntity == "%({DOWN})") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.DOWN); //System.Windows.Forms.Keys.Alt);

                return;
            }
            if (myEntity == "%({F8})") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.F8); //System.Windows.Forms.Keys.Alt);
                return;
            }
            if (myEntity == "{NUMPADADD}") {
                InputSimulator.Keyboard.KeyPress(VirtualKeyCode.ADD);
                return;
            }
            if (myEntity == "{NUMPADMULT}") {
                InputSimulator.Keyboard.KeyPress(VirtualKeyCode.MULTIPLY);
                return;
            }
            if (myEntity == "{ENTER}") {
                InputSimulator.Keyboard.KeyPress(VirtualKeyCode.RETURN);
                return;
            }
            if (myEntity == "^(n)") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_N);
                return;
            }
            if (myEntity == "^(v)") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_V);
                return;
            }
            if (myEntity == "^(c)") {
                try {

                    Thread thread = new Thread(new ThreadStart(() => {
                        Clipboard.Clear();
                        // Clipboard.SetText(myEntity);            
                        // or call logic here



                    }));

                    thread.SetApartmentState(ApartmentState.STA);

                    thread.Start();
                    thread.Join();
                } catch (Exception ex) {
                    MessageBox.Show("Here is an exception thrown in TypeText method in IdealAutomateCore for myEntity " + myEntity + ": " + ex.Message);
                }
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.VK_C);
                return;
            }
            if (myEntity == "{F5}") {
                InputSimulator.Keyboard.KeyPress(VirtualKeyCode.F5);
                return;
            }
            if (myEntity == "{F6}") {
                InputSimulator.Keyboard.KeyPress(VirtualKeyCode.F6);
                return;
            }
            if (myEntity == "{DOWN}") {
                InputSimulator.Keyboard.KeyPress(VirtualKeyCode.DOWN);
                return;
            }
            if (myEntity == "{UP}") {
                InputSimulator.Keyboard.KeyPress(VirtualKeyCode.UP);
                return;
            }
            if (myEntity == "^({END})") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.END);
                return;
            }
            if (myEntity == "^({HOME})") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.CONTROL, VirtualKeyCode.HOME);
                return;
            }
            if (myEntity == "+({F10})") {
                InputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.SHIFT, VirtualKeyCode.F10);
                return;
            }


            System.Windows.Forms.SendKeys.SendWait(myEntity);

        }
        /// <summary>
        /// PutClipboardInEntity returns a string that contains the text in the clipboard.
        /// </summary>
        /// <returns>string that contains the text in the clipboard</returns>

        public string PutClipboardInEntity() {
            if (fbDebugMode) {
                Console.WriteLine(oProcess.ProcessName + "==> " + "PutClipboardInEntity: ");

            }
            int intTryAgainCtr = 0;
            TryAgain:
            string myEntity = "";
            try {
                //Thread thread = new Thread(new ThreadStart(() => {
                try {

                    for (int i = 0; i < 45; i++) {
                        if (Clipboard.GetData(DataFormats.Text) == null) {
                            Sleep(100);
                        } else {
                            break;
                        }
                        if (i == 10 || i == 20 || i == 30 || i == 40) {
                            TypeText("^(c)", 500);
                            Sleep(1001);
                            // System.Diagnostics.Debugger.Break();
                        }
                    }

                    myEntity = Clipboard.GetData(DataFormats.Text).ToString();
                    if (myEntity.Contains("Failed to copy selection to the clipboard")) {
                        TypeText("^(c)", 500);
                        Sleep(1002);
                        intTryAgainCtr++;
                        if (intTryAgainCtr < 10) {
                            goto TryAgain;
                        }

                    }
                    // myEntity = Clipboard.GetText(System.Windows.TextDataFormat.Html);
                } catch (Exception e) {
                    Console.WriteLine("Exception occurred in PutClipboardInEntity!!!! " + e.Message);
                    myEntity = "";
                }

                // or call logic here



                //}));

                //thread.SetApartmentState(ApartmentState.STA);

                //thread.Start();
                //thread.Join();
            } catch (Exception ex) {
                try {
                    TypeText("^(c)", 500);
                    //Thread thread = new Thread(new ThreadStart(() => {
                    try {

                        for (int i = 0; i < 45; i++) {
                            if (Clipboard.GetData(DataFormats.Text) == null) {
                                Sleep(100);
                            } else {
                                break;
                            }
                            if (i == 10 || i == 20 || i == 30 || i == 40) {
                                TypeText("^(c)", 500);
                                Sleep(1003);
                            }
                        }

                        myEntity = Clipboard.GetData(DataFormats.Text).ToString();
                        // myEntity = Clipboard.GetText(System.Windows.TextDataFormat.Html);
                    } catch (Exception e) {
                        Console.WriteLine("Exception occurred in PutClipboardInEntity!!!! " + e.Message);
                        myEntity = "";
                    }

                    // or call logic here



                    //}));

                    //thread.SetApartmentState(ApartmentState.STA);

                    //thread.Start();
                    //thread.Join();
                } catch (Exception ex1) {
                    MessageBox.Show(ex1.Message);
                }
            }

            Console.Write("myEntity=" + myEntity);


            return myEntity;
        }
        /// <summary>
        /// Sleep receives an integer that indicates the number of milliseconds that you want the program to wait.
        /// </summary>
        /// <param name="intSleep">integer that indicates the number of milliseconds that you want the program to wait.</param>
        public void Sleep(int intSleep) {
            if (fbDebugMode) {
                Console.WriteLine(oProcess.ProcessName + "==> " + "Sleep:  intSleep=" + intSleep.ToString());

            }
            System.Threading.Thread.Sleep(intSleep);
        }
        public void FindDelimitedText(FindDelimitedTextParms myParms) {
            // set defaults of not found and empty string
            myParms.intDelimFound = -1;
            myParms.strDelimitedTextFound = "";
            int intStartingLine = myParms.intLineCtr;

            ExamineALine:
            // if we reach the end of the lines array,
            // we could not find delimited text and we are done
            if (myParms.intLineCtr > myParms.lines.Count()) {
                myParms.strResultTypeFound = "Not Found";
                return;
            }
            // move the current line that we want to examine
            // to myLine

            string myLine = myParms.lines[myParms.intLineCtr];
            if (myParms.intLineCtr == intStartingLine) {
                myLine = myLine.Substring(myParms.intStartingCol);
            }

            // we are going to loop the the begin and end 
            // delim pairs looking for a match for any pair
            // in the entire lines array starting at the initial
            // intLinesCtr.
            // this will exit on the first one that it finds

            for (int i = 0; i < myParms.lsBeginDelim.Count; i++) {
                string strBeginDelim = myParms.lsBeginDelim[i];
                string strEndDelim = myParms.lsEndDelim[i];
                // Find location of beginning delim
                int intBeginDelimLength = strBeginDelim.Length;

                LookForBeginDelim:
                int indexBeginDelim = myLine.ToUpper().IndexOf(strBeginDelim);
                // Get the rest of the line after Begin Delim
                if (indexBeginDelim == -1) {
                    if (myParms.intLineCtr < myParms.lines.Count()) {
                        myParms.intLineCtr++;
                        myLine = myParms.lines[myParms.intLineCtr];
                        goto LookForBeginDelim;
                    }
                }

                // if we went thru the entire lines array
                // and did not find BeginDelim, we need to
                // go get the next begin and end delim pair
                // and start over

                if (indexBeginDelim == -1) {
                    myParms.intLineCtr = intStartingLine;
                    myLine = myParms.lines[myParms.intLineCtr];
                    myLine = myLine.Substring(myParms.intStartingCol);
                    continue;
                }

                // we have found the begin delim so we want to save
                // it as strResultType to tell us which begin and end
                // delim pair we found. We still need to look for the 
                // end delim in order to have a successful find

                strBeginDelim = myLine.Substring(indexBeginDelim, intBeginDelimLength);
                if (myParms.lsBeginDelim.Count > 1) {
                    myParms.strResultTypeFound = strBeginDelim;
                }
                myParms.intDelimFound = i;
                string d = myLine.Substring(indexBeginDelim + intBeginDelimLength);
                // Find location of EndDelim

                LookForEndDelim:

                int indexEndDelim = d.IndexOf(strEndDelim);
                if (indexEndDelim == -1) {
                    if (myParms.intLineCtr < myParms.lines.Count()) {
                        myParms.intLineCtr++;
                        myLine = myParms.lines[myParms.intLineCtr];
                        goto LookForEndDelim;
                    }
                }

                // if we went thru the entire lines array
                // and did not find EndDelim, we need to
                // go get the next begin and end delim pair
                // and start over

                if (indexEndDelim == -1) {
                    myParms.intLineCtr = intStartingLine;
                    myLine = myParms.lines[myParms.intLineCtr];
                    myLine = myLine.Substring(myParms.intStartingCol);
                    continue;
                }

                d = d.Substring(0, indexEndDelim);
                myParms.strDelimitedTextFound = d;
                myParms.intEndDelimColPosFound = indexEndDelim + strEndDelim.Length;
                return;
            }

            if (myParms.lsBeginDelim.Count > 1) {
                myParms.intLineCtr++;
                goto ExamineALine;
            }
        }
        /// <summary>
        /// <para>GetAppDirectoryForScript gets the application </para>
        /// <para>data folder and adds \IdealAutomate\yourscriptname to it.</para>
        /// <para>The AppDirectory allows you to store personal settings and</para>
        /// <para>information that you want to keep private (like passwords) in a location</para>
        /// <para>outside of your script on in the application directory</para>
        /// </summary>
        /// <returns>string that is the app_data/roaming directory path for the script</returns>
        public string GetAppDirectoryForScript() {

            string directory = AppDomain.CurrentDomain.BaseDirectory;
            directory = directory.Replace("\\bin\\Debug\\", "");
            int intLastSlashIndex = directory.LastIndexOf("\\");
            //string strScriptName = directory.Substring(intLastSlashIndex + 1);
            // string strScriptName = System.Reflection.Assembly.GetCallingAssembly().GetName().Name;
            string settingsDirectory =
      Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\IdealAutomate\\" + ConvertFullFileNameToScriptPath(directory);
            if (!Directory.Exists(settingsDirectory)) {
                Directory.CreateDirectory(settingsDirectory);
            }
            return settingsDirectory;
        }

        /// <summary>
        /// <para>GetAppDirectoryForIdealAutomate gets the application </para>
        /// <para>data folder and adds \IdealAutomate\yourscriptname to it.</para>
        /// <para>The AppDirectory allows you to store personal settings and</para>
        /// <para>information that you want to keep private (like passwords) in a location</para>
        /// <para>outside of your script on in the application directory</para>
        /// </summary>
        /// <returns>string that is the app_data/roaming directory path for the script</returns>
        public string GetAppDirectoryForIdealAutomate() {
            string directory = AppDomain.CurrentDomain.BaseDirectory;
            directory = directory.Replace("\\bin\\Debug\\", "");
            int intLastSlashIndex = directory.LastIndexOf("\\");
            string strScriptName = directory.Substring(intLastSlashIndex + 1);
            // string strScriptName = System.Reflection.Assembly.GetCallingAssembly().GetName().Name;
            string settingsDirectory =
      Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\IdealAutomate";
            if (!Directory.Exists(settingsDirectory)) {
                Directory.CreateDirectory(settingsDirectory);
            }
            return settingsDirectory;
        }


        /// <summary>
        /// <para>ReadValueFromAppDataKey takes a key and adds .txt to it in order to create</para>
        /// <para>a file name. It gets the app data path and adds \IdealAutomate\yourscriptname</para>
        /// <para>to it. By combining that path to the file name created from the key,</para>
        /// <para>it can retrieve a value from the key that is unique to your script application.</para>
        ///  <para>The AppDirectory allows you to store personal settings and</para>
        /// <para>information that you want to keep private (like passwords) in a location</para>
        /// <para>outside of your script on in the application directory</para>
        /// </summary>
        /// <param name="strKey">Unique key within the script application</param>
        /// <returns>string that was in application directory for that key</returns>
        public string GetValueByKey(string strKey) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = GetAppDirectoryForScript();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            return strValueRead;
        }

        public int GetValueByKeyAsInt(string strKey) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = GetAppDirectoryForScript();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            int intValue = 0;
            Int32.TryParse(strValueRead, out intValue);
            return intValue;
        }

        public int GetValueByKeyAsIntForNonCurrentScript(string strKey, string strScriptName) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = GetAppDirectoryForIdealAutomate();
            settingsDirectory = Path.Combine(settingsDirectory, strScriptName);
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            int intValue = 0;
            Int32.TryParse(strValueRead, out intValue);
            return intValue;
        }

        public int GetValueByPublicKeyAsIntForNonCurrentScript(string strKey, string strFullFolderPath) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = strFullFolderPath;
            settingsDirectory = Path.Combine(settingsDirectory, "..IdealAutomate");
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            int intValue = 0;
            Int32.TryParse(strValueRead, out intValue);
            return intValue;
        }

        public string GetValueByKeyForNonCurrentScript(string strKey, string strScriptName) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = GetAppDirectoryForIdealAutomate();
            settingsDirectory = Path.Combine(settingsDirectory, strScriptName);
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            return strValueRead;
        }

        public string GetValueByPublicKeyForNonCurrentScript(string strKey, string strFullFolderPath) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = strFullFolderPath;
            settingsDirectory = Path.Combine(settingsDirectory, "..IdealAutomate");
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            return strValueRead;
        }

        public double GetValueByKeyAsDouble(string strKey) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = GetAppDirectoryForScript();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            double intValue = 0;
            double.TryParse(strValueRead, out intValue);
            return intValue;
        }
        public DateTime GetValueByKeyAsDateTime(string strKey) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = GetAppDirectoryForScript();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            DateTime dtScriptStartDateTime = DateTime.MinValue;
            bool boolDTSuccessful = DateTime.TryParse(strValueRead, out dtScriptStartDateTime);
            return dtScriptStartDateTime;
        }

        public DateTime GetValueByKeyAsDateTimeForNonCurrentScript(string strKey, string strScriptName) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = GetAppDirectoryForIdealAutomate();
            settingsDirectory = Path.Combine(settingsDirectory, strScriptName);
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            DateTime dtScriptStartDateTime = DateTime.MinValue;
            bool boolDTSuccessful = DateTime.TryParse(strValueRead, out dtScriptStartDateTime);
            return dtScriptStartDateTime;
        }

        public DateTime GetValueByPublicKeyAsDateTimeForNonCurrentScript(string strKey, string strFullFolderPath) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = strFullFolderPath;
            settingsDirectory = Path.Combine(settingsDirectory, "..IdealAutomate");
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            DateTime dtScriptStartDateTime = DateTime.MinValue;
            bool boolDTSuccessful = DateTime.TryParse(strValueRead, out dtScriptStartDateTime);
            return dtScriptStartDateTime;
        }

        public int GetValueByKeyAsIntGlobal(string strKey) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = GetAppDirectoryForIdealAutomate();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            int intValue = 0;
            Int32.TryParse(strValueRead, out intValue);
            return intValue;
        }

        public int IncrementValueByKeyByValue(string strKey, int intIncrementValue) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = GetAppDirectoryForScript();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            int intValue = 0;
            Int32.TryParse(strValueRead, out intValue);
            int intResult = intValue + intIncrementValue;
            SetValueByKey(strKey, intResult.ToString());
            return intResult;
        }

        public int IncrementValueByKeyByValueGlobal(string strKey, int intIncrementValue) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = GetAppDirectoryForIdealAutomate();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            int intValue = 0;
            Int32.TryParse(strValueRead, out intValue);
            int intResult = intValue + intIncrementValue;
            SetValueByKey(strKey, intResult.ToString());
            return intResult;
        }

        /// <summary>
        /// <para>SetValueByKey takes a key and adds .txt to it in order to create</para>
        /// <para>a file name. It gets the app data path and adds \IdealAutomate\yourscriptname</para>
        /// <para>to it. By combining that path to the file name,</para>
        /// <para>it can write a value to the key that is unique to your script application.</para>
        /// <para>The AppDirectory allows you to store personal settings and</para>
        /// <para>information that you want to keep private (like passwords) in a location</para>
        /// <para>outside of your script on in the application directory</para>
        /// </summary>
        /// <param name="strKey">Unique key within the script application</param>
        /// <param name="strValueToWrite">Value to write to the Unique key 
        /// within the script application</param>
        public void SetValueByKey(string strKey, string strValueToWrite) {
            string fileName = strKey + ".txt";
            StreamWriter writer = null;
            string settingsDirectory = GetAppDirectoryForScript();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            // Hook a write to the text file.
            writer = new StreamWriter(settingsPath);
            // Rewrite the entire value of s to the file
            writer.Write(strValueToWrite);
            writer.Close();
        }

        public void SetValueByKeyForNonCurrentScript(string strKey, string strValueToWrite, string strScriptName) {
            string fileName = strKey + ".txt";
            StreamWriter writer = null;
            string settingsDirectory = GetAppDirectoryForIdealAutomate();
            settingsDirectory = Path.Combine(settingsDirectory, strScriptName);
            if (!Directory.Exists(settingsDirectory)) {
                Directory.CreateDirectory(settingsDirectory);
            }
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            // Hook a write to the text file.
            writer = new StreamWriter(settingsPath);
            // Rewrite the entire value of s to the file
            writer.Write(strValueToWrite);
            writer.Close();
        }

        public void SetValueByPublicKeyForNonCurrentScript(string strKey, string strValueToWrite, string strFullFolderPath) {
            string fileName = strKey + ".txt";
            StreamWriter writer = null;
            string settingsDirectory = strFullFolderPath;
            settingsDirectory = Path.Combine(settingsDirectory, "..IdealAutomate");
            if (!Directory.Exists(settingsDirectory)) {
                Directory.CreateDirectory(settingsDirectory);
            }
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            // Hook a write to the text file.
            writer = new StreamWriter(settingsPath);
            // Rewrite the entire value of s to the file
            writer.Write(strValueToWrite);
            writer.Close();
        }

        /// <summary>
        /// <para>GetValueByKeyGlobal takes a key and adds .txt to it in order to create</para>
        /// <para>a file name. It gets the app data path and adds \IdealAutomate\yourscriptname</para>
        /// <para>to it. By combining that path to the file name created from the key,</para>
        /// <para>it can retrieve a value from the key that is unique to your script application.</para>
        ///  <para>The AppDirectory allows you to store personal settings and</para>
        /// <para>information that you want to keep private (like passwords) in a location</para>
        /// <para>outside of your script on in the application directory</para>
        /// </summary>
        /// <param name="strKey">Unique key within the script application</param>
        /// <returns>string that was in application directory for that key</returns>
        public string GetValueByKeyGlobal(string strKey) {
            string fileName = strKey + ".txt";
            StreamReader file = null;
            string strValueRead = "";
            string settingsDirectory = GetAppDirectoryForIdealAutomate();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                file = File.OpenText(settingsPath);
                strValueRead = file.ReadToEnd();
                file.Close();
            }
            return strValueRead;
        }

        /// <summary>
        /// <para>SetValueByKeyGlobal takes a key and adds .txt to it in order to create</para>
        /// <para>a file name. It gets the app data path and adds \IdealAutomate\yourscriptname</para>
        /// <para>to it. By combining that path to the file name,</para>
        /// <para>it can write a value to the key that is unique to your script application.</para>
        /// <para>The AppDirectory allows you to store personal settings and</para>
        /// <para>information that you want to keep private (like passwords) in a location</para>
        /// <para>outside of your script on in the application directory</para>
        /// </summary>
        /// <param name="strKey">Unique key within the script application</param>
        /// <param name="strValueToWrite">Value to write to the Unique key 
        /// within the script application</param>
        public void SetValueByKeyGlobal(string strKey, string strValueToWrite) {
            string fileName = strKey + ".txt";
            StreamWriter writer = null;
            string settingsDirectory = GetAppDirectoryForIdealAutomate();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            // Hook a write to the text file.
            writer = new StreamWriter(settingsPath);
            // Rewrite the entire value of s to the file
            writer.Write(strValueToWrite);
            writer.Close();
        }

        /// <summary>
        /// <para>ReadAppDirectoryKeyToArrayList takes a key and adds .txt to it in order to create</para>
        /// <para>a file name. It gets the app data path and adds \IdealAutomate\yourscriptname</para>
        /// <para>to it. By combining that path to the file name created from the key,</para>
        /// <para>it can retrieve an arraylist that is unique to your script application.</para>
        /// <para>The AppDirectory allows you to store personal settings and</para>
        /// <para>information that you want to keep private (like passwords) in a location</para>
        /// <para>outside of your script on in the application directory</para>
        /// </summary>
        /// <param name="strKey">Unique key within the script application</param>
        /// <returns>ArrayList that was in application directory for that key</returns>
        public ArrayList ReadAppDirectoryKeyToArrayList(string strKey) {
            string fileName = strKey + ".txt";
            ArrayList myArrayList = new ArrayList();
            string settingsDirectory = GetAppDirectoryForScript();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                try {
                    StreamReader reader = File.OpenText(settingsPath);
                    while (!reader.EndOfStream) {
                        string myLine = reader.ReadLine();
                        myArrayList.Add(myLine);
                    }
                    reader.Close();
                } catch (Exception ex) {

                    string message = "Error - Reading  " + fileName + " " + ex.Message + " " + ex.InnerException; // +"; EntityName is: " + myEntityForExecutable.EntityName;
                    MessageBox.Show(message);
                }
            }
            return myArrayList;
        }

        /// <summary>
        /// <para>WriteArrayListToAppDirectoryKey takes a key and adds .txt to it in order to create</para>
        /// <para>a file name. It gets the app data path and adds \IdealAutomate\yourscriptname</para>
        /// <para>to it. By combining that path to the file name created from the key,</para>
        /// <para>it can write an arraylist to the key filename is unique to your script application.</para>
        /// <para>The AppDirectory allows you to store personal settings and</para>
        /// <para>information that you want to keep private (like passwords) in a location</para>
        /// <para>outside of your script on in the application directory</para>
        /// </summary>
        /// <param name="strKey">Unique key within the script application</param>
        /// <param name="arrayListToWrite">ArrayList that is to written to the application directory for that key</param>
        public void WriteArrayListToAppDirectoryKey(string strKey, ArrayList arrayListToWrite) {
            string fileName = strKey + ".txt";
            StreamWriter writer = null;
            string settingsDirectory = GetAppDirectoryForScript();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            // Hook a write to the text file.
            writer = new StreamWriter(settingsPath);
            // Rewrite the entire value of s to the file
            foreach (var item in arrayListToWrite) {
                writer.WriteLine(item.ToString());
            }
            writer.Close();
        }

        /// <summary>
        /// <para>ReadAppDirectoryKeyToArrayList takes a key and adds .txt to it in order to create</para>
        /// <para>a file name. It gets the app data path and adds \IdealAutomate\yourscriptname</para>
        /// <para>to it. By combining that path to the file name created from the key,</para>
        /// <para>it can retrieve an arraylist that is unique to your script application.</para>
        /// <para>The AppDirectory allows you to store personal settings and</para>
        /// <para>information that you want to keep private (like passwords) in a location</para>
        /// <para>outside of your script on in the application directory</para>
        /// </summary>
        /// <param name="strKey">Unique key within the script application</param>
        /// <returns>ArrayList that was in application directory for that key</returns>
        public ArrayList ReadAppDirectoryKeyToArrayListGlobal(string strKey) {
            string fileName = strKey + ".txt";
            ArrayList myArrayList = new ArrayList();
            string settingsDirectory = GetAppDirectoryForIdealAutomate();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            if (File.Exists(settingsPath)) {
                try {
                    StreamReader reader = File.OpenText(settingsPath);
                    while (!reader.EndOfStream) {
                        string myLine = reader.ReadLine();
                        myArrayList.Add(myLine);
                    }
                    reader.Close();
                } catch (Exception ex) {

                    string message = "Error - Reading  " + fileName + " " + ex.Message + " " + ex.InnerException; // +"; EntityName is: " + myEntityForExecutable.EntityName;
                    MessageBox.Show(message);
                }
            }
            return myArrayList;
        }

        /// <summary>
        /// <para>WriteArrayListToAppDirectoryKey takes a key and adds .txt to it in order to create</para>
        /// <para>a file name. It gets the app data path and adds \IdealAutomate\yourscriptname</para>
        /// <para>to it. By combining that path to the file name created from the key,</para>
        /// <para>it can write an arraylist to the key filename is unique to your script application.</para>
        /// <para>The AppDirectory allows you to store personal settings and</para>
        /// <para>information that you want to keep private (like passwords) in a location</para>
        /// <para>outside of your script on in the application directory</para>
        /// </summary>
        /// <param name="strKey">Unique key within the script application</param>
        /// <param name="arrayListToWrite">ArrayList that is to written to the application directory for that key</param>
        public void WriteArrayListToAppDirectoryKeyGlobal(string strKey, ArrayList arrayListToWrite) {
            string fileName = strKey + ".txt";
            StreamWriter writer = null;
            string settingsDirectory = GetAppDirectoryForIdealAutomate();
            string settingsPath = Path.Combine(settingsDirectory, fileName);
            // Hook a write to the text file.
            writer = new StreamWriter(settingsPath);
            // Rewrite the entire value of s to the file
            foreach (var item in arrayListToWrite) {
                writer.WriteLine(item.ToString());
            }
            writer.Close();
        }
        /// <summary>
        /// At beginning of template, Write StartTime, StartDate, Add 1 to total executions. 
        /// 
        /// </summary>
        public void ScriptStartedUpdateStats() {
            SetValueByKey("ScriptStartDateTime", System.DateTime.Now.ToString());
            IncrementValueByKeyByValue("ScriptTotalExecutions", 1);

        }

        /// <summary>
        /// ScriptEndedSuccessfullyUpdateStats - At end of template, Subtract start time and start date
        /// from Current date and time, add 1 to successful executions. 
        /// (avg successful execution time * 
        /// (total successful executions - 1)) + current successful execution time) / total successful executions. 
        /// Write PercentSuccessful = (SuccessfulExecutions / TotalExecutions) * 100. 
        /// CurrentSavedExecutionTime = CurrentExecutionTime - ManualExecutionTime.
        /// ScriptTotalSavedExecutionTime += CurrentSavedExecutionTime
        ///  StartDate is the LastExecutedDate.
        /// </summary>
        public void ScriptEndedSuccessfullyUpdateStats() {
            DateTime dtScriptStartDateTime = GetValueByKeyAsDateTime("ScriptStartDateTime");


            if (dtScriptStartDateTime == DateTime.MinValue) {
                return;
            }

            TimeSpan ts = System.DateTime.Now - dtScriptStartDateTime;

            int intAvgSuccessfulExecutionTime = GetValueByKeyAsInt("AvgSuccessfulExecutionTime");
            int intScriptSuccessfulExecutions = GetValueByKeyAsInt("ScriptSuccessfulExecutions");
            int intTotalSuccessfulExecutionTimePrev = intAvgSuccessfulExecutionTime *
                (intScriptSuccessfulExecutions);
            int intTotalSuccessfulExecutionTime = intTotalSuccessfulExecutionTimePrev + (int)ts.TotalSeconds;
            intScriptSuccessfulExecutions = IncrementValueByKeyByValue("ScriptSuccessfulExecutions", 1);
            intAvgSuccessfulExecutionTime = intTotalSuccessfulExecutionTime / intScriptSuccessfulExecutions;
            SetValueByKey("AvgSuccessfulExecutionTime", intAvgSuccessfulExecutionTime.ToString());
            int intScriptTotalExecutions = GetValueByKeyAsInt("ScriptTotalExecutions");
            decimal decPercentSuccessful = (intScriptSuccessfulExecutions / intScriptTotalExecutions) * 100;
            int intPercentSuccessful = (int)Math.Round(decPercentSuccessful);
            SetValueByKey("ScriptPercentSuccessful", intPercentSuccessful.ToString());
            int intManualExecutionTime = GetValueByKeyAsInt("ManualExecutionTime");
            int intScriptCurrentExecutionTimeSavings = 0;
            if (intManualExecutionTime != 0) {
                intScriptCurrentExecutionTimeSavings = intManualExecutionTime - intTotalSuccessfulExecutionTime;
            }

            IncrementValueByKeyByValue("ScriptTotalSavedExecutionTime", intScriptCurrentExecutionTimeSavings);




        }
        public string ConvertFullFileNameToScriptPath(string fullFileName) {
            int intIndex = fullFileName.LastIndexOf(@"\");
            if (intIndex > -1) {
                fullFileName = fullFileName.Substring(0, intIndex);
            }
            string scriptPath = fullFileName;
            scriptPath = scriptPath.Replace(":", "+").Replace(@"\", "-");
            return scriptPath;
        }

        public string ConvertFullFileNameToPublicPath(string fullFileName) {
            int intIndex = fullFileName.LastIndexOf(@"\");
            if (intIndex > -1) {
                fullFileName = fullFileName.Substring(0, intIndex);
            }
            string scriptPath = fullFileName;
            return scriptPath;
        }

        public string ConvertFullFileNameToScriptPathWithoutRemoveLastLevel(string fullFileName) {
            string scriptPath = fullFileName;
            scriptPath = scriptPath.Replace(":", "+").Replace(@"\", "-");
            return scriptPath;
        }

        public string GetPathForScriptNoBinDebug() {
            string directory = AppDomain.CurrentDomain.BaseDirectory;
            directory = directory.Replace("\\bin\\Debug\\", "");
            return directory;

        }

    }

   
}
