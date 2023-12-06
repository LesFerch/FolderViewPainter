using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Windows.Automation;
using System.Diagnostics;
using System.Threading;
using System.Drawing;
using System.Reflection;

namespace FolderViewPainter
{
    class Program
    {
        static string myName = typeof(Program).Namespace;
        static string ShellKey = @"Software\Classes\Local Settings\Software\Microsoft\Windows\Shell";
        static string ShellBagsKey = $@"{ShellKey}\Bags";
        static string ShellBagsNet = @"Software\Microsoft\Windows\Shell\Bags";

        static string sMain = "Folder View Painter";
        static string[] MenuKeys = { "Import", "Export", "Manage", "Options", "Help" };
        static string[] cmds = { "/i", "/e", "/m", "/o", "/h" };
        static string sMenuLabels = "Import View|Export View|Manage|Options|Help";
        static string[] MenuLabels = sMenuLabels.Split(new char[] { '|' });
        static string sOK = "OK";
        static string sYes = "Yes";
        static string sNo = "No";
        static string sInstall = "Install";
        static string sRemove = "Remove";
        static string sDone = "Done";
        static string sInput = "Enter a name for this saved view";
        static string sCount1Msg = "Registry views total";
        static string sCount2Msg = "Exported views total";
        static string sNoRoot = "Root folders not supported";
        static string sAlreadyExists = "Name already exists. Overwrite?";
        static string sSetup = "Install or Remove this tool";
        static string sImportInterface = "Import Interface";
        static string sExportInterface = "Export Interface";
        static string sQuickpickdialog = "Quick pick dialog";
        static string sQuicksavedialog = "Quick save dialog";
        static string sFilesavedialog = "File save dialog";
        static string sFileopendialog = "File open dialog";
        static string sNoExportedViews = "No exported views found";

        static string Folder = "";
        static string MyFolder = AppDomain.CurrentDomain.BaseDirectory;
        static string RegFolder = $@"{MyFolder}SavedViews\";
        static string RegFile = "";
        static string newFileName = "";
        static string NodeType1 = null;
        static string NodeType2 = null;
        static string NodeType3 = null;
        static string NodeType4 = null;
        static string ExportNode;
        static string LeafPath;
        static string ThisPCPath;
        static string UserPath;
        static string[] NodeList;
        static string[] GUIDList;
        static float ScaleFactor;
        static int dialogX = 0;
        static int dialogY = 0;
        static bool Dark = isDark();
        static bool useOpenDialog = false;
        static bool useSaveDialog = false;

        [STAThread]
        static void Main(string[] args)
        {
            LoadLanguageStrings();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            ScaleFactor = GetScale();
            LoadSettingsFromRegistry();

            if (args.Length == 0)
            {
                InstallRemove();
                return;
            }

            Directory.CreateDirectory(RegFolder);
            string command = "";

            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].StartsWith("/")) { command = args[i].ToLower(); }
                else
                {
                    Folder = args[i].Replace("\"", "").Trim();
                }
            }

            if (command == "/install") { InstallContextMenuEntries(); return; }

            if (command == "/remove") { RemoveContextMenuEntries(); return; }

            if (command == "/h") { Process.Start($@"https://lesferch.github.io/{myName}/"); return; }

            if (command == "/o")
            {
                var optionsDialog = new OptionsDialog();
                optionsDialog.ShowDialog();
                return;
            }

            if (command == "/m") { Process.Start($"explorer.exe", $"\"{RegFolder}\""); return; }

            if (Folder == "") { return; }

            if (Folder.StartsWith("\\")) { ShellBagsKey = ShellBagsNet; }

            LeafPath = Path.GetFileName(Folder.TrimEnd('\\'));
            if (LeafPath == "") { LeafPath = Folder[0].ToString(); }


            string UserName = Environment.GetEnvironmentVariable("UserName");
            UserPath = Folder.Replace($@"C:\Users\{UserName}", UserName);
            ThisPCPath = Folder.Replace($@"C:\Users\{UserName}", "This PC");

            bool IsSpecial = IsShellFolder(Folder);
            bool SaveView = false;

            if (command == "/e")
            {
                if (useSaveDialog)
                {
                    string exeDirectory = Path.GetDirectoryName(Application.ExecutablePath);
                    string savedViewsFolder = Path.Combine(exeDirectory, "SavedViews");

                    if (Directory.Exists(savedViewsFolder))
                    {
                        SaveFileDialog fd = new SaveFileDialog
                        {
                            Filter = "Reg files (*.reg)|*.reg",
                            InitialDirectory = savedViewsFolder,
                            FileName = LeafPath
                        };
                        DialogResult result = fd.ShowDialog();
                        if (result == DialogResult.OK)
                        {
                            RegFile = fd.FileName;
                            if (RegFile == "") { return; }
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                else
                {
                    while (true)
                    {
                        SaveView = true;
                        string sPrompt = $"{sCount1Msg} = {GetNodeSlotsCount()}\n{sCount2Msg} = {GetSavedViewsCount()}\n\n{sInput}";

                        newFileName = CustomInputDialog.ShowDialog(sPrompt, sMain, LeafPath);
                        if (newFileName == "") { return; }

                        RegFile = $@"{RegFolder}{newFileName}.reg";
                        if (File.Exists(RegFile))
                        {
                            DialogResult result = TwoChoiceBox.Show($"{sAlreadyExists}", sMain, sYes, sNo);
                            if (result == DialogResult.Yes) { break; }
                            if (result == DialogResult.Cancel) { return; }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
            }

            if (command == "/i")
            {
                if (useOpenDialog)
                {
                    string exeDirectory = Path.GetDirectoryName(Application.ExecutablePath);
                    string savedViewsFolder = Path.Combine(exeDirectory, "SavedViews");

                    if (Directory.Exists(savedViewsFolder))
                    {
                        OpenFileDialog fd = new OpenFileDialog
                        {
                            Filter = "Reg files (*.reg)|*.reg",
                            InitialDirectory = savedViewsFolder,
                            Multiselect = false,
                        };
                        fd.ShowDialog();
                        RegFile = fd.FileName;
                        if (RegFile == "") { return; }
                    }
                }
                else
                {
                    string pickedItem = ShowPickListDialog();
                    if (pickedItem == null) { return; }
                    RegFile = $@"{RegFolder}{pickedItem}.reg";
                }
            }

            GetBagNodes();

            for (int i = 0; i < NodeList.Length; i++)
            {
                DeleteRevValue(NodeList[i], GUIDList[i]);
            }

            CloseWindows(Folder, LeafPath);

            Thread.Sleep(500);

            GetBagNodes();

            ExportNode = NodeList[0];

            for (int i = 0; i < NodeList.Length; i++)
            {
                if (ExistRevValue(NodeList[i], GUIDList[i])) { ExportNode = NodeList[i]; }
            }

            if ((NodeType1 == null) && IsSpecial)
            {
                Process.Start("explorer.exe", Folder);
                Thread.Sleep(600);

                CloseWindows(Folder, LeafPath);
                Thread.Sleep(500);

                GetBagNodes();
            }

            if (SaveView)
            {
                RegFile = $@"{RegFolder}{newFileName}.reg";
                string GUID = GetGUID($@"{ShellBagsKey}\{ExportNode}\Shell\");
                string keyPath = $@"HKCU\{ShellBagsKey}\{ExportNode}\Shell\{GUID}";
                RunProcess("reg.exe", $"export \"{keyPath}\" \"{RegFile}\" /y");
            }

            for (int i = 0; i < NodeList.Length; i++)
            {
                UpdateRegFile(NodeList[i], GUIDList[i], RegFile);
                RunProcess("reg.exe", $"import \"{RegFile}\"");
            }

            Process.Start("explorer.exe", Folder);
        }

        //--------------------------------------------------------------------------------------------------------------

        [Flags]
        public enum SIGDN : uint
        {
            NORMALDISPLAY = 0x00000000,
            FILESYSPATH = 0x80058000,
        }

        // Load language strings from INI file
        static void LoadLanguageStrings()
        {
            string lang = GetLang();

            IniFile iniFile = new IniFile("language.ini");

            sMenuLabels = iniFile.ReadString(lang, "sMenuLabels", sMenuLabels);
            MenuLabels = sMenuLabels.Split(new char[] { '|' });

            sMain = iniFile.ReadString(lang, "sMain", sMain);
            sOK = iniFile.ReadString(lang, "sOK", sOK);
            sYes = iniFile.ReadString(lang, "sYes", sYes);
            sNo = iniFile.ReadString(lang, "sNo", sNo);
            sInstall = iniFile.ReadString(lang, "sInstall", sInstall);
            sRemove = iniFile.ReadString(lang, "sRemove", sRemove);
            sDone = iniFile.ReadString(lang, "sDone", sDone);
            sInput = iniFile.ReadString(lang, "sInput", sInput);
            sCount1Msg = iniFile.ReadString(lang, "sCount1Msg", sCount1Msg);
            sCount2Msg = iniFile.ReadString(lang, "sCount2Msg", sCount2Msg);
            sNoRoot = iniFile.ReadString(lang, "sNoRoot", sNoRoot);
            sAlreadyExists = iniFile.ReadString(lang, "sAlreadyExists", sAlreadyExists);
            sSetup = iniFile.ReadString(lang, "sSetup", sSetup);
            sImportInterface = iniFile.ReadString(lang, "sImportInterface", sImportInterface);
            sExportInterface = iniFile.ReadString(lang, "sExportInterface", sExportInterface);
            sQuickpickdialog = iniFile.ReadString(lang, "sQuickpickdialog", sQuickpickdialog);
            sQuicksavedialog = iniFile.ReadString(lang, "sQuicksavedialog", sQuicksavedialog);
            sFilesavedialog = iniFile.ReadString(lang, "sFilesavedialog", sFilesavedialog);
            sFileopendialog = iniFile.ReadString(lang, "sFileopendialog", sFileopendialog);
            sNoExportedViews = iniFile.ReadString(lang, "sNoExportedViews", sNoExportedViews);
        }

        // INI file parser
        public class IniFile
        {
            private readonly string iniFilePath;

            public IniFile(string fileName)
            {
                string exePath = Assembly.GetExecutingAssembly().Location;
                string exeDirectory = System.IO.Path.GetDirectoryName(exePath);
                iniFilePath = System.IO.Path.Combine(exeDirectory, fileName);
            }

            public string ReadString(string section, string key, string defaultValue)
            {
                try
                {
                    if (File.Exists(iniFilePath))
                    {
                        return IniFileParser.ReadValue(section, key, defaultValue, iniFilePath);
                    }
                }
                catch { }

                return defaultValue;
            }
        }

        public static class IniFileParser
        {
            public static string ReadValue(string section, string key, string defaultValue, string filePath)
            {
                try
                {
                    var lines = File.ReadAllLines(filePath, Encoding.UTF8);
                    string currentSection = null;

                    foreach (var line in lines)
                    {
                        string trimmedLine = line.Trim();

                        if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]"))
                        {
                            currentSection = trimmedLine.Substring(1, trimmedLine.Length - 2);
                        }
                        else if (currentSection == section)
                        {
                            var parts = trimmedLine.Split(new char[] { '=' }, 2);
                            if (parts.Length == 2 && parts[0].Trim() == key)
                            {
                                return parts[1].Trim();
                            }
                        }
                    }
                }
                catch (Exception)
                {
                }
                return defaultValue;
            }
        }


        // Get the current system language
        static string GetLang()
        {
            string lang = "en";
            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\International");
                if (key != null)
                {
                    lang = key.GetValue("LocaleName") as string;
                    key.Close();
                }
            }
            catch { }

            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey("Control Panel\\Desktop");
                if (key != null)
                {
                    string[] preferredLanguages = key.GetValue("PreferredUILanguages") as string[];
                    if (preferredLanguages != null && preferredLanguages.Length > 0)
                    {
                        lang = preferredLanguages[0];
                    }
                    key.Close();
                }
            }
            catch { }

            return lang.Substring(0, 2).ToLower();
        }

        // Convert BagMRU entry IDL to name
        public class IDL2Name
        {
            [DllImport("shell32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
            [return: MarshalAs(UnmanagedType.Error)]
            public static extern int SHGetNameFromIDList(IntPtr pidl, SIGDN sigdnName, out StringBuilder ppszName);

            public static string Get_FolderName(byte[] IDL, SIGDN sigdnName)
            {
                GCHandle pinnedArray = GCHandle.Alloc(IDL, GCHandleType.Pinned);
                IntPtr PIDL = pinnedArray.AddrOfPinnedObject();
                StringBuilder name = new StringBuilder(2048);
                int result = SHGetNameFromIDList(PIDL, sigdnName, out name);
                pinnedArray.Free();
                if (result == 0)
                {
                    return name.ToString();
                }
                else
                {
                    return null;
                }
            }
        }

        // Decode BagMRU entries to a lits of paths and nodes
        public class BagMRUEntry
        {
            public string BagNode { get; set; }
            public string BagPath { get; set; }
        }

        public static List<BagMRUEntry> GetBagMRU(RegistryKey[] keys = null, byte[] IDL = null, string parentPath = null, byte[] parentIDL = null)
        {
            var result = new List<BagMRUEntry>();
            string KeyLoc = @"Software\Classes\Local Settings\Software\Microsoft\Windows\Shell\BagMRU";
            string KeyNet = @"Software\Microsoft\Windows\Shell\BagMRU";

            if (keys == null)
            {
                keys = new[]
                {
                    Registry.CurrentUser.OpenSubKey(KeyLoc),
                    Registry.CurrentUser.OpenSubKey(KeyNet)
                };
            }

            if (IDL == null)
            {
                IDL = new byte[0];
            }

            if (parentIDL == null)
            {
                parentIDL = new byte[0];
            }

            foreach (var key in keys)
            {
                string nameSpacePath = null, fileSysPath = null;

                if (IDL.Length > 0)
                {
                    byte[] absIDL = parentIDL.Concat(IDL).ToArray();

                    try
                    {
                        nameSpacePath = parentPath + IDL2Name.Get_FolderName(absIDL, SIGDN.NORMALDISPLAY);
                    }
                    catch
                    {
                        nameSpacePath = null;
                    }

                    try
                    {
                        fileSysPath = IDL2Name.Get_FolderName(absIDL, SIGDN.FILESYSPATH);
                    }
                    catch
                    {
                        fileSysPath = null;
                    }
                }

                string bagNum = key.GetValue("NodeSlot")?.ToString() ?? null;

                if ((bagNum != null) && (fileSysPath != null))
                {
                    if (!nameSpacePath.Contains('(')) { fileSysPath = nameSpacePath; }
                    if (key.Name.Contains("Classes") && fileSysPath.StartsWith(@"\\")) { }
                    else { result.Add(new BagMRUEntry { BagNode = bagNum, BagPath = fileSysPath }); }
                }
                if (IDL.Length > 0)
                {
                    parentPath = $"{nameSpacePath}\\";
                    parentIDL = parentIDL.Concat(IDL.Take(IDL.Length - 2)).ToArray();
                }
                else
                {
                    parentPath = "";
                    parentIDL = new byte[0];
                }

                foreach (string subKeyName in key.GetSubKeyNames())
                {
                    var subKey = key.OpenSubKey(subKeyName);
                    byte[] subKeyIDL = (byte[])key.GetValue(subKeyName);
                    var subResult = GetBagMRU(new[] { subKey }, subKeyIDL, parentPath, parentIDL);
                    result.AddRange(subResult);
                }
            }

            return result;
        }

        // Put nodes in a dictionary keyed by path
        static void GetBagNodes()
        {
            var result = GetBagMRU();

            Dictionary<string, string> BagDictionary;

            BagDictionary = result
            .GroupBy(item => item.BagPath)
            .ToDictionary(
                group => group.Key,
                group => group
                    .Select(item => item.BagNode)
                    .Aggregate((current, next) => current + "|" + next));

            BagDictionary.TryGetValue(Folder, out NodeType1);
            BagDictionary.TryGetValue(ThisPCPath, out NodeType2);
            BagDictionary.TryGetValue(UserPath, out NodeType3);
            BagDictionary.TryGetValue(LeafPath, out NodeType4);

            string Nodes = $"{NodeType1}|{NodeType2}|{NodeType3}|{NodeType4}";

            NodeList = Nodes.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            GUIDList = Nodes.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);

            for (int i = 0; i < NodeList.Length; i++)
            {
                GUIDList[i] = GetGUID($@"{ShellBagsKey}\{NodeList[i]}\Shell\");
            }

        }

        // Determine if the current folder is a shell folder
        static bool IsShellFolder(string folderPath)
        {
            const string userShellFoldersPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders";

            if (folderPath.StartsWith(@"C:\Users")) { return true; }

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(userShellFoldersPath))
            {
                if (key != null)
                {
                    foreach (var valueName in key.GetValueNames())
                    {
                        // Get the value and expand any environment variables
                        string registryPath = key.GetValue(valueName).ToString();
                        string expandedPath = Environment.ExpandEnvironmentVariables(registryPath);

                        //Get the length of the shortest of the two paths to compare
                        int n = Math.Min(expandedPath.Length, folderPath.Length);
                        string path1 = expandedPath.Substring(0, n).ToLower();
                        string path2 = folderPath.Substring(0, n).ToLower();

                        // Check if the expanded path matches the input folder path
                        if (path1 == path2)
                        {
                            return true;
                        }
                    }
                }
            }
            // Return false if the path was not found
            return false;
        }

        // Find the matching Explorer window, so it can be closed
        static AutomationElement FindWindowByName(string name)
        {
            return AutomationElement.RootElement.FindFirst(
                TreeScope.Children,
                new PropertyCondition(AutomationElement.NameProperty, name)
            );
        }

        // Find root view window by partial match to drive letter with parentheses
        static AutomationElement FindWindowByPName(string name)
        {
            var windows = AutomationElement.RootElement.FindAll(
                TreeScope.Children,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
            );

            foreach (AutomationElement window in windows)
            {
                string windowName = window.Current.Name;
                if (windowName.Contains($"({name.Replace("\\","")})"))
                {
                    return window;
                }
            }

            return null;
        }

        // Close the found window
        static void CloseWindow(AutomationElement window)
        {
            try
            {
                // Attempt to find the close button and invoke it
                AutomationElement closeButton = window.FindFirst(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, "Close"));
                if (closeButton != null)
                {
                    InvokePattern invokePattern = closeButton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    invokePattern?.Invoke();
                }
            }
            catch
            {
            }
        }

        // Get the GUID subkey of the selected node key
        static string GetGUID(string key)
        {
            string GUID = "";
            try
            {
                using (RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(key))
                {
                    if (registryKey != null)
                    {
                        string[] subKeyNames = registryKey.GetSubKeyNames();
                        if (subKeyNames.Length > 0) { GUID = subKeyNames[0]; }
                    }
                }
            }
            catch
            {
            }
            return GUID;
        }

        // Update exported view reg file with the target (import) node and GUID 
        static void UpdateRegFile(string NodeNum, string GUID, string RegFile)
        {
            string newline = $@"[HKEY_CURRENT_USER\{ShellBagsKey}\{NodeNum}\Shell\{GUID}]";
            try
            {
                string[] lines = File.ReadAllLines(RegFile);

                for (int i = 0; i < lines.Length; i++)
                {
                    if (lines[i].TrimStart().StartsWith("["))
                    {
                        lines[i] = newline;
                        break;
                    }
                }
                File.WriteAllLines(RegFile, lines);
            }
            catch
            {
            }
        }

        static void InstallRemove()
        {
            DialogResult result = TwoChoiceBox.Show(sSetup, sMain, sInstall, sRemove);

            if (result == DialogResult.Yes)
            {
                InstallContextMenuEntries();
                CustomMessageBox.Show(sDone, sMain);
            }
            if (result == DialogResult.No)
            {
                RemoveContextMenuEntries();
                CustomMessageBox.Show(sDone, sMain);
            }
        }

        static void InstallContextMenuEntries()
        {

            RemoveContextMenuEntries();

            string exePath = System.Reflection.Assembly.GetExecutingAssembly().Location;

            string MyKey = $@"Software\Classes\Directory\background\shell\{myName}";

            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(MyKey))
            {
                key.SetValue("SubCommands", "");
                key.SetValue("", "");
                key.SetValue("MUIVerb", sMain);
                key.SetValue("Icon", exePath);
            }

            for (int i = 0; i < MenuLabels.Length; i++)
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey($@"{MyKey}\shell\{i}-{MenuKeys[i]}"))
                {
                    key.SetValue("", MenuLabels[i]);

                    using (RegistryKey commandKey = key.CreateSubKey("command"))
                    {
                        string cmd = $"\"{exePath}\" {cmds[i]}";
                        if (i < 2) { cmd += " %v"; }
                        commandKey.SetValue("", cmd);
                    }
                }
            }
        }

        static void RemoveContextMenuEntries()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Classes\Directory\background\shell", true))
            {
                try { key.DeleteSubKeyTree($@"{myName}", false); }
                catch { }
            }
        }

        // Close Explorer windows that match the curent path
        static void CloseWindows(string Folder, string LeafPath)
        {
            AutomationElement foundWindow;

            do
            {
                foundWindow = FindWindowByName(Folder);
                if (foundWindow == null) { break; }
                CloseWindow(foundWindow);
                Thread.Sleep(10);
            }
            while (foundWindow != null);

            if (Folder.Length < 4)
            {
                do
                {
                    foundWindow = FindWindowByPName(Folder);
                    if (foundWindow == null) { break; }
                    CloseWindow(foundWindow);
                    Thread.Sleep(10);
                }
                while (foundWindow != null);
            }
            else
            {
                do
                {
                    foundWindow = FindWindowByName(LeafPath);
                    if (foundWindow == null) { break; }
                    CloseWindow(foundWindow);
                    Thread.Sleep(10);
                }
                while (foundWindow != null);
            }
        }

        // Trick for determining which namespace view we're looking at
        // Delete the "Rev" key for all nodes with a matching path
        // The one that gets recreated when we close the current window
        // is the correct node for the export
        static void DeleteRevValue(string NodeNum, string GUID)
        {
            string keyPath = $@"{ShellBagsKey}\{NodeNum}\Shell\{GUID}";

            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath, true))
                {
                    if (key != null && key.GetValue("Rev") != null) { key.DeleteValue("Rev"); }
                }
            }
            catch (Exception)
            {
            }
        }

        // Check if "Rev" key exists
        static bool ExistRevValue(string NodeNum, string GUID)
        {
            string keyPath = $@"HKEY_CURRENT_USER\{ShellBagsKey}\{NodeNum}\Shell\{GUID}";
            return Registry.GetValue(keyPath, "Rev", null) != null;
        }

        static void RunProcess(string executablePath, string arguments)
        {
            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = executablePath,
                Arguments = arguments,
                CreateNoWindow = true,
                UseShellExecute = false
            };
            using (Process process = new Process { StartInfo = psi })
            {
                process.Start();
                process.WaitForExit();
            }
        }

        // Get count of exported view reg files
        static int GetSavedViewsCount()
        {
            string exeDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            string savedViewsFolder = Path.Combine(exeDirectory, "SavedViews");
            string[] regFiles = Directory.GetFiles(savedViewsFolder, "*.reg");
            return regFiles.Length;
        }

        // Get count of saved vies in registry (aka nodes).
        static int GetNodeSlotsCount()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey($@"{ShellKey}\BagMRU"))
            {
                if (key != null)
                {
                    byte[] nodeSlots = key.GetValue("NodeSlots") as byte[];
                    if (nodeSlots != null)
                    {
                        return nodeSlots.Length;
                    }
                }
            }
            return -1;
        }

        // Dialog for selecting a view to import
        static string ShowPickListDialog()
        {
            using (var dialog = new CustomForm())
            {
                dialog.Font = new Font("Segoe UI", 10);

                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.ControlBox = true;
                dialog.MaximizeBox = false;
                dialog.MinimizeBox = false;
                dialog.StartPosition = FormStartPosition.Manual;
                dialog.ShowInTaskbar = false;
                dialog.AutoSize = true;
                dialog.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                dialog.AutoScroll = true;


                if (Dark)
                {
                    DarkTitleBar(dialog.Handle);
                    dialog.BackColor = Color.FromArgb(32, 32, 32);
                    dialog.ForeColor = Color.White;
                }

                string exeDirectory = Path.GetDirectoryName(Application.ExecutablePath);
                string savedViewsFolder = Path.Combine(exeDirectory, "SavedViews");

                if (Directory.Exists(savedViewsFolder))
                {
                    string[] regFiles = Directory.GetFiles(savedViewsFolder, "*.reg");
                    if (regFiles.Length == 0)
                    {
                        CustomMessageBox.Show($"{sNoExportedViews}", sMain);
                        return null;
                    }

                    int yOffset = 0;
                    int maxWidth = 0;
                    int minWidth = (int)(120 * ScaleFactor);

                    using (Graphics g = dialog.CreateGraphics())
                    {
                        foreach (string regFile in regFiles)
                        {
                            string itemName = Path.GetFileNameWithoutExtension(regFile);
                            SizeF size = g.MeasureString(itemName, new Font("Segoe UI", 10));
                            maxWidth = Math.Max(maxWidth, (int)size.Width);
                        }
                    }
                    maxWidth += (int)(20 * ScaleFactor);
                    if (maxWidth < minWidth) { maxWidth = minWidth; }

                    foreach (string regFile in regFiles)
                    {
                        string itemName = Path.GetFileNameWithoutExtension(regFile);

                        var label = new Label
                        {
                            Text = itemName,
                            Cursor = Cursors.Hand,
                            Width = maxWidth,
                            Height = (int)(20 * ScaleFactor),
                            Location = new Point((int)(4 * ScaleFactor), yOffset),
                            AutoSize = false,
                        };

                        label.MouseEnter += (sender, e) =>
                        {
                            if (Dark)
                            {
                                label.BackColor = label.BackColor = Color.FromArgb(0, 120, 215);
                            }
                            else
                            {
                                label.BackColor = label.BackColor = Color.FromArgb(145, 201, 247);
                            }
                        };

                        label.MouseLeave += (sender, e) =>
                        {
                            label.BackColor = dialog.BackColor;
                        };

                        label.Click += (sender, e) =>
                        {
                            dialog.Tag = itemName;
                            dialog.DialogResult = DialogResult.OK;
                            dialog.Close();
                        };

                        dialog.Controls.Add(label);

                        yOffset += label.Height;
                    }
                }

                Point cursorPosition = Cursor.Position;

                Screen screen = Screen.FromPoint(cursorPosition);

                int screenWidth = screen.WorkingArea.Width;
                int screenHeight = screen.WorkingArea.Height;

                if (dialog.Height > screenHeight)
                {
                    dialog.Height = screenHeight;
                    dialog.Width += (int)(16 * ScaleFactor); //increase width for scrollbar
                    dialog.AutoSize = false;
                }

                dialogX = Cursor.Position.X - dialog.Width / 2;
                dialogY = Cursor.Position.Y - dialog.Height / 2;

                int baseX = screen.Bounds.X;
                int baseY = screen.Bounds.Y;

                dialogX = Math.Max(baseX, Math.Min(baseX + screenWidth - dialog.Width, dialogX));
                dialogY = Math.Max(baseY, Math.Min(baseY + screenHeight - dialog.Height, dialogY));

                dialog.Location = new Point(dialogX, dialogY);

                DialogResult result = dialog.ShowDialog();

                if (result == DialogResult.OK && dialog.Tag != null)
                {
                    return dialog.Tag.ToString();
                }
                else
                {
                    return null;
                }
            }
        }

        // Get current screen scaling factor
        static float GetScale()
        {
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                float dpiX = graphics.DpiX;
                return dpiX / 96;
            }
        }

        // Remove the horizontal scroll bar from the Import dialog
        public class CustomForm : Form
        {
            private const int WS_HSCROLL = 0x00100000;
            private const int WM_NCCALCSIZE = 0x0083;

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_NCCALCSIZE)
                {
                    int style = GetWindowLong(Handle, GWL_STYLE);
                    if ((style & WS_HSCROLL) == WS_HSCROLL)
                    {
                        style &= ~WS_HSCROLL;
                        SetWindowLong(Handle, GWL_STYLE, style);
                    }
                }
                base.WndProc(ref m);
            }

            [DllImport("user32.dll", SetLastError = true)]
            private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

            [DllImport("user32.dll", SetLastError = true)]
            private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

            private const int GWL_STYLE = -16;
        }

        // Dialog for exporting a view
        public class CustomInputDialog : Form
        {
            private TextBox textBoxInput;
            private Button buttonOK;
            private Label labelPrompt;

            public CustomInputDialog(string prompt, string title, string defaultResponse)
            {
                Text = title;
                Font = new Font("Segoe UI", 10);
                StartPosition = FormStartPosition.Manual;
                Width = (int)(420 * ScaleFactor);
                Height = (int)(210 * ScaleFactor);
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                Move += new EventHandler(Form_Move);

                labelPrompt = new Label();
                labelPrompt.Text = prompt;
                labelPrompt.Left = (int)(10 * ScaleFactor);
                labelPrompt.Top = (int)(10 * ScaleFactor);
                labelPrompt.Width = ClientSize.Width - (int)(20 * ScaleFactor);
                labelPrompt.AutoSize = true;
                Controls.Add(labelPrompt);

                textBoxInput = new TextBox();
                textBoxInput.Left = (int)(10 * ScaleFactor);
                textBoxInput.Top = labelPrompt.Bottom + (int)(10 * ScaleFactor);
                textBoxInput.Width = ClientSize.Width - (int)(20 * ScaleFactor);
                textBoxInput.Text = defaultResponse;
                textBoxInput.BorderStyle = BorderStyle.FixedSingle;
                Controls.Add(textBoxInput);

                buttonOK = new Button();
                buttonOK.Text = sOK;
                buttonOK.Width = (int)(75 * ScaleFactor);
                buttonOK.Height = (int)(26 * ScaleFactor);
                buttonOK.Left = (ClientSize.Width - buttonOK.Width) / 2;
                buttonOK.Top = ClientSize.Height - buttonOK.Height - (int)(10 * ScaleFactor);
                buttonOK.Click += new EventHandler(buttonOK_Click);
                buttonOK.Font = new Font("Segoe UI", 9);
                if (Dark)
                {
                    buttonOK.FlatStyle = FlatStyle.Flat;
                    buttonOK.FlatAppearance.BorderColor = SystemColors.Highlight;
                    buttonOK.FlatAppearance.BorderSize = 1;
                    DarkTitleBar(Handle);
                    BackColor = Color.FromArgb(32, 32, 32);
                    ForeColor = Color.White;
                    textBoxInput.BackColor = Color.FromArgb(32, 32, 32);
                    textBoxInput.ForeColor = Color.White;
                }
                Controls.Add(buttonOK);
                AcceptButton = buttonOK;

                if ((dialogX == 0) && (dialogY == 0))
                {
                    Point cursorPosition = Cursor.Position;
                    dialogX = Cursor.Position.X - Width / 2;
                    dialogY = Cursor.Position.Y - Height / 2;
                    Screen screen = Screen.FromPoint(cursorPosition);
                    int screenWidth = screen.WorkingArea.Width;
                    int screenHeight = screen.WorkingArea.Height;
                    int baseX = screen.Bounds.X;
                    int baseY = screen.Bounds.Y;
                    dialogX = Math.Max(baseX, Math.Min(baseX + screenWidth - Width, dialogX));
                    dialogY = Math.Max(baseY, Math.Min(baseY + screenHeight - Height, dialogY));
                }
                Location = new Point(dialogX, dialogY);
            }
            private void Form_Move(object sender, EventArgs e)
            {
                dialogX = Location.X;
                dialogY = Location.Y;
            }
            private void buttonOK_Click(object sender, EventArgs e)
            {
                DialogResult = DialogResult.OK;
                Close();
            }
            public static string ShowDialog(string prompt, string title, string defaultResponse)
            {
                using (CustomInputDialog inputDialog = new CustomInputDialog(prompt, title, defaultResponse))
                {
                    var result = inputDialog.ShowDialog();
                    return result == DialogResult.OK ? inputDialog.textBoxInput.Text : "";
                }
            }
        }

        // Dialog for simple OK messages
        public class CustomMessageBox : Form
        {
            private Label messageLabel;
            private Button buttonOK;

            public CustomMessageBox(string message, string caption)
            {
                message = $"\n{message}";

                StartPosition = FormStartPosition.Manual;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                Text = caption;
                Width = (int)(300 * ScaleFactor);
                Height = (int)(150 * ScaleFactor);
                MaximizeBox = false;
                MinimizeBox = false;

                messageLabel = new Label();
                messageLabel.Text = message;
                messageLabel.Font = new Font("Segoe UI", 10);
                messageLabel.TextAlign = ContentAlignment.TopCenter;
                messageLabel.Dock = DockStyle.Fill;

                using (Graphics g = CreateGraphics())
                {
                    SizeF size = g.MeasureString(message, new Font("Segoe UI", 10), Width);
                    Height = Math.Max(Height, (int)(size.Height + (int)(100 * ScaleFactor)));
                }


                buttonOK = new Button();
                buttonOK.Text = sOK;
                buttonOK.DialogResult = DialogResult.OK;
                buttonOK.Font = new Font("Segoe UI", 9);
                buttonOK.Width = (int)(75 * ScaleFactor);
                buttonOK.Height = (int)(26 * ScaleFactor);
                buttonOK.Left = (ClientSize.Width - buttonOK.Width) / 2;
                buttonOK.Top = ClientSize.Height - buttonOK.Height - (int)(10 * ScaleFactor);
                if (Dark)
                {
                    buttonOK.FlatStyle = FlatStyle.Flat;
                    buttonOK.FlatAppearance.BorderColor = SystemColors.Highlight;
                    buttonOK.FlatAppearance.BorderSize = 1;
                    DarkTitleBar(Handle);
                    BackColor = Color.FromArgb(32, 32, 32);
                    ForeColor = Color.White;
                }
                Controls.Add(buttonOK);
                Controls.Add(messageLabel);

                Point cursorPosition = Cursor.Position;
                int dialogX = Cursor.Position.X - Width / 2;
                int dialogY = Cursor.Position.Y - Height / 2;
                Screen screen = Screen.FromPoint(cursorPosition);
                int screenWidth = screen.WorkingArea.Width;
                int screenHeight = screen.WorkingArea.Height;
                int baseX = screen.Bounds.X;
                int baseY = screen.Bounds.Y;
                dialogX = Math.Max(baseX, Math.Min(baseX + screenWidth - Width, dialogX));
                dialogY = Math.Max(baseY, Math.Min(baseY + screenHeight - Height, dialogY));
                Location = new Point(dialogX, dialogY);
            }
            public static DialogResult Show(string message, string caption)
            {
                using (var customMessageBox = new CustomMessageBox(message, caption))
                {
                    return customMessageBox.ShowDialog();
                }
            }

        }

        // Dialog for install/Remove and Yes/No
        public class TwoChoiceBox : Form
        {
            private Label messageLabel;
            private Button buttonYes;
            private Button buttonNo;

            public TwoChoiceBox(string message, string caption, string button1, string button2)
            {
                int b2Width = (int)(75 * ScaleFactor);
                using (Graphics g = CreateGraphics())
                {
                    SizeF size = g.MeasureString(button2, new Font("Segoe UI", 9));
                    b2Width = Math.Max((int)size.Width, b2Width);
                }
                message = $"\n{message}";

                StartPosition = FormStartPosition.Manual;
                Text = caption;
                Width = (int)(300 * ScaleFactor);
                Height = (int)(150 * ScaleFactor);
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;

                messageLabel = new Label();
                messageLabel.Text = message;
                messageLabel.Font = new Font("Segoe UI", 10);
                messageLabel.TextAlign = ContentAlignment.TopCenter;
                messageLabel.Dock = DockStyle.Fill;

                buttonYes = new Button();
                buttonYes.Text = button1;
                buttonYes.Font = new Font("Segoe UI", 9);
                buttonYes.DialogResult = DialogResult.OK;
                buttonYes.MinimumSize = new Size((int)(75 * ScaleFactor), (int)(26 * ScaleFactor));
                buttonYes.Left = (int)(10 * ScaleFactor);
                buttonYes.Top = ClientSize.Height - buttonYes.Height - (int)(12 * ScaleFactor);
                buttonYes.DialogResult = DialogResult.Yes;

                buttonNo = new Button();
                buttonNo.Text = button2;
                buttonNo.Font = new Font("Segoe UI", 9);
                buttonNo.DialogResult = DialogResult.OK;
                buttonNo.MinimumSize = new Size((int)(75 * ScaleFactor), (int)(26 * ScaleFactor));
                buttonNo.Left = ClientSize.Width - b2Width - (int)(16 * ScaleFactor);
                buttonNo.Top = ClientSize.Height - buttonNo.Height - (int)(12 * ScaleFactor);
                buttonNo.DialogResult = DialogResult.No;

                if (Dark)
                {
                    buttonYes.FlatStyle = FlatStyle.Flat;
                    buttonYes.FlatAppearance.BorderColor = SystemColors.Highlight;
                    buttonYes.FlatAppearance.BorderSize = 1;
                    buttonNo.FlatStyle = FlatStyle.Flat;
                    buttonNo.FlatAppearance.BorderColor = SystemColors.Highlight;
                    buttonNo.FlatAppearance.BorderSize = 1;
                    DarkTitleBar(Handle);
                    BackColor = Color.FromArgb(32, 32, 32);
                    ForeColor = Color.White;
                }

                Controls.Add(buttonYes);
                Controls.Add(buttonNo);
                Controls.Add(messageLabel);

                Point cursorPosition = Cursor.Position;
                int dialogX = Cursor.Position.X - Width / 2;
                int dialogY = Cursor.Position.Y - Height / 2;
                Screen screen = Screen.FromPoint(cursorPosition);
                int screenWidth = screen.WorkingArea.Width;
                int screenHeight = screen.WorkingArea.Height;
                int baseX = screen.Bounds.X;
                int baseY = screen.Bounds.Y;
                dialogX = Math.Max(baseX, Math.Min(baseX + screenWidth - Width, dialogX));
                dialogY = Math.Max(baseY, Math.Min(baseY + screenHeight - Height, dialogY));
                Location = new Point(dialogX, dialogY);
            }

            public static DialogResult Show(string message, string caption, string button1, string button2)
            {
                using (var TwoChoiceBox = new TwoChoiceBox(message, caption, button1, button2))
                {
                    return TwoChoiceBox.ShowDialog();
                }
            }

        }

        // Determine if dark colors (theme) are being used
        public static bool isDark()
        {
            const string keyPath = @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize";
            const string valueName = "AppsUseLightTheme";

            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath))
            {
                if (key != null)
                {
                    object value = key.GetValue(valueName);
                    if (value is int intValue)
                    {
                        return intValue == 0;
                    }
                }
            }
            return false; // Return false if the key or value is missing
        }

        // Dialog for selecting Options
        public class OptionsDialog : Form
        {
            private RadioButton quickSaveRadioButton;
            private RadioButton classicSaveRadioButton;
            private RadioButton quickOpenRadioButton;
            private RadioButton classicOpenRadioButton;
            private Button buttonOK;

            public OptionsDialog()
            {
                InitializeComponents();
                LoadSettingsFromRegistry();
                classicSaveRadioButton.Checked = useSaveDialog;
                quickSaveRadioButton.Checked = !useSaveDialog;
                classicOpenRadioButton.Checked = useOpenDialog;
                quickOpenRadioButton.Checked = !useOpenDialog;
            }

            private void InitializeComponents()
            {
                Text = sMain;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                ShowIcon = false;
                StartPosition = FormStartPosition.Manual;
                Width = (int)(300 * ScaleFactor);
                Height = (int)(270 * ScaleFactor);
                Font = new Font("Segoe UI", 10);

                var importGroupBox = new GroupBox
                {
                    Name = "importGroupBox",
                    Text = sImportInterface,
                    Location = new Point((int)(10 * ScaleFactor), (int)(10 * ScaleFactor)),
                    Width = Width - (int)(40 * ScaleFactor),
                    Height = (int)(80 * ScaleFactor)
                };

                if (Dark) { importGroupBox.ForeColor = Color.White; }

                quickOpenRadioButton = new RadioButton
                {
                    Text = sQuickpickdialog,
                    Location = new Point((int)(10 * ScaleFactor), (int)(20 * ScaleFactor)),
                    AutoSize = true,
                    Checked = true
                };

                classicOpenRadioButton = new RadioButton
                {
                    Text = sFileopendialog,
                    Location = new Point((int)(10 * ScaleFactor), (int)(40 * ScaleFactor)),
                    AutoSize = true
                };

                importGroupBox.Controls.Add(quickOpenRadioButton);
                importGroupBox.Controls.Add(classicOpenRadioButton);

                var exportGroupBox = new GroupBox
                {
                    Name = "exportGroupBox",
                    Text = sExportInterface,
                    Left = (int)(10 * ScaleFactor),
                    Top = (int)(100 * ScaleFactor),
                    Width = Width - (int)(40 * ScaleFactor),
                    Height = (int)(80 * ScaleFactor)
                };

                if (Dark) { exportGroupBox.ForeColor = Color.White; }

                quickSaveRadioButton = new RadioButton
                {
                    Text = sQuicksavedialog,
                    Left = (int)(10 * ScaleFactor),
                    Top = (int)(20 * ScaleFactor),
                    AutoSize = true,
                    Checked = true
                };

                classicSaveRadioButton = new RadioButton
                {
                    Text = sFilesavedialog,
                    Left = (int)(10 * ScaleFactor),
                    Top = (int)(40 * ScaleFactor),
                    AutoSize = true
                };

                exportGroupBox.Controls.Add(quickSaveRadioButton);
                exportGroupBox.Controls.Add(classicSaveRadioButton);

                if (Dark)
                {
                    DarkTitleBar(Handle);
                    BackColor = Color.FromArgb(32, 32, 32);
                    ForeColor = Color.White;
                }

                Controls.Add(importGroupBox);
                Controls.Add(exportGroupBox);

                buttonOK = new Button();
                buttonOK.Text = sOK;
                buttonOK.Left = (Width - (int)(80 * ScaleFactor)) / 2;
                buttonOK.Top = Height - (int)(74 * ScaleFactor);
                buttonOK.Width = (int)(75 * ScaleFactor);
                buttonOK.Height = (int)(26 * ScaleFactor);
                buttonOK.Click += new EventHandler(buttonOK_Click);
                buttonOK.Font = new Font("Segoe UI", 9);
                if (Dark)
                {
                    buttonOK.FlatStyle = FlatStyle.Flat;
                    buttonOK.FlatAppearance.BorderColor = SystemColors.Highlight;
                    buttonOK.FlatAppearance.BorderSize = 1;
                }
                Controls.Add(buttonOK);

                Point cursorPosition = Cursor.Position;
                int dialogX = Cursor.Position.X - Width / 2;
                int dialogY = Cursor.Position.Y - Height / 2;
                Screen screen = Screen.FromPoint(cursorPosition);
                int screenWidth = screen.WorkingArea.Width;
                int screenHeight = screen.WorkingArea.Height;
                int baseX = screen.Bounds.X;
                int baseY = screen.Bounds.Y;
                dialogX = Math.Max(baseX, Math.Min(baseX + screenWidth - Width, dialogX));
                dialogY = Math.Max(baseY, Math.Min(baseY + screenHeight - Height, dialogY));
                Location = new Point(dialogX, dialogY);
            }

            private void buttonOK_Click(object sender, EventArgs e)
            {
                SaveSettingsToRegistry();
                Close();
            }

            private void SaveSettingsToRegistry()
            {
                using (var key = Registry.CurrentUser.CreateSubKey(@"Software\FolderViewPainter"))
                {
                    if (key != null)
                    {
                        key.SetValue("useSaveDialog", classicSaveRadioButton.Checked ? 1 : 0, RegistryValueKind.DWord);
                        key.SetValue("useOpenDialog", classicOpenRadioButton.Checked ? 1 : 0, RegistryValueKind.DWord);
                    }
                }
            }
        }

        // Load current option settings from registry
        static void LoadSettingsFromRegistry()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(@"Software\FolderViewPainter"))
            {
                if (key != null)
                {
                    useSaveDialog = Convert.ToBoolean(key.GetValue("useSaveDialog", 0));
                    useOpenDialog = Convert.ToBoolean(key.GetValue("useOpenDialog", 0));
                }
            }
        }

        // Make dialog title bar black
        public enum DWMWINDOWATTRIBUTE : uint
        {
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20,
        }

        [DllImport("dwmapi.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
        public static extern void DwmSetWindowAttribute(IntPtr hwnd, DWMWINDOWATTRIBUTE attribute, ref int pvAttribute, uint cbAttribute);

        static void DarkTitleBar(IntPtr hWnd)
        {
            var preference = Convert.ToInt32(true);
            DwmSetWindowAttribute(hWnd, DWMWINDOWATTRIBUTE.DWMWA_USE_IMMERSIVE_DARK_MODE, ref preference, sizeof(uint));

        }
    }
}