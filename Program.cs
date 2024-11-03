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
        static string myExe = Assembly.GetExecutingAssembly().Location;
        static string ShellKey = @"Software\Classes\Local Settings\Software\Microsoft\Windows\Shell";
        static string ShellBagsKey = $@"{ShellKey}\Bags";
        static string ShellBagsNet = @"Software\Microsoft\Windows\Shell\Bags";

        static string sMain = "Folder View Painter";
        static string[] MenuKeys = { "Import", "Export", "Reset", "Manage", "Options", "Help" };
        static string[] cmds = { "/i", "/e", "/r", "/m", "/o", "/h" };
        static string sMenuLabels = "Import View|Export View|Reset View|Manage|Options|Help";
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
        static string sIncludeSettings = "Include Explorer settings";
        static string sResetViewSetup = "Include additional Explorer settings with Reset View.";
        static string sResetViewNote = "Recheck to update saved settings to current values.";

        static string Folder = "";
        static string MyFolder = AppDomain.CurrentDomain.BaseDirectory;
        static string RegFolder = $@"{MyFolder}SavedViews\";
        static string ResetViewFile = $@"{RegFolder}!ResetView.ini";
        static string myIniFile = $@"{MyFolder}\{myName}.ini";
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

        static CheckBox checkboxExp;

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

            if (command == "/h") { Process.Start($@"https://lesferch.github.io/{myName}/#how-to-use"); return; }

            if (!CanWriteHere())
            {
                RegFolder = $@"{Environment.GetEnvironmentVariable("AppData")}\SavedViews\";
                ResetViewFile = $@"{RegFolder}!ResetView.ini";
            }
            Directory.CreateDirectory(RegFolder);

            if (command == "/o")
            {
                var optionsDialog = new OptionsDialog();
                optionsDialog.ShowDialog();
                return;
            }

            if (command == "/m") { Process.Start("explorer.exe", $"\"{RegFolder}\""); return; }

            if (Folder == "") { return; }

            if (Folder.StartsWith("\\")) { ShellBagsKey = ShellBagsNet; }

            LeafPath = Path.GetFileName(Folder.TrimEnd('\\'));
            if (LeafPath == "") { LeafPath = Folder[0].ToString(); }

            string UserProfile = Environment.GetEnvironmentVariable("UserProfile");
            string ProfileDir = Path.GetFileName(UserProfile);
            UserPath = Folder.Replace(UserProfile, ProfileDir);
            ThisPCPath = Folder.Replace(UserProfile, "This PC");

            bool IsSpecial = IsShellFolder(Folder);
            bool SaveView = false;

            if (command == "/e")
            {
                if (useSaveDialog)
                {
                    if (Directory.Exists(RegFolder))
                    {
                        SaveFileDialog fd = new SaveFileDialog
                        {
                            Filter = "Reg files (*.reg)|*.reg",
                            InitialDirectory = RegFolder,
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
                            DialogResult result = TwoChoiceBox.Show(sAlreadyExists, sMain, sYes, sNo);
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
                    if (Directory.Exists(RegFolder))
                    {
                        OpenFileDialog fd = new OpenFileDialog
                        {
                            Filter = "Reg files (*.reg)|*.reg",
                            InitialDirectory = RegFolder,
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
                ExportRegistryKey(keyPath, RegFile, true);
                if (checkboxExp.Checked) ExportExplorerSettings(RegFile, false);
            }

            for (int i = 0; i < NodeList.Length; i++)
            {
                if (command == "/r")
                {
                    DeleteAllValues(NodeList[i], GUIDList[i]);
                }
                else
                {
                    UpdateRegFile(NodeList[i], GUIDList[i], RegFile);
                    ImportRegistryFile(RegFile);
                }
            }

            if (command == "/r" && File.Exists(ResetViewFile)) ImportRegistryFile(ResetViewFile);

            Process.Start("explorer.exe", Folder);
        }

        static void ExportExplorerSettings(string RegFile, bool overWrite)
        {
            string keyPath = $@"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Modules\GlobalSettings";
            ExportRegistryKey(keyPath, RegFile, overWrite);
            keyPath = $@"HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced";
            ExportRegistryKey(keyPath, RegFile, false);
        }

        //--------------------------------------------------------------------------------------------------------------

        [Flags]
        public enum SIGDN : uint
        {
            NORMALDISPLAY = 0x00000000,
            FILESYSPATH = 0x80058000,
        }

        // Test if app can write to its own folder (portable install)
        static bool CanWriteHere()
        {
            try
            {
                string tempFile = $"{MyFolder}temp.txt";
                File.WriteAllText(tempFile, "");
                File.Delete(tempFile);
                return true;
            }
            catch { return false; }
        }

        // Load language strings from INI file
        static void LoadLanguageStrings()
        {
            string iniFile = $@"{MyFolder}\language.ini";

            if (!File.Exists(iniFile)) return;

            string lang = GetLang();

            sMenuLabels = ReadString(iniFile, lang, "sMenuLabels", sMenuLabels);
            MenuLabels = sMenuLabels.Split(new char[] { '|' });

            sMain = ReadString(iniFile, lang, "sMain", sMain);
            sOK = ReadString(iniFile, lang, "sOK", sOK);
            sYes = ReadString(iniFile, lang, "sYes", sYes);
            sNo = ReadString(iniFile, lang, "sNo", sNo);
            sInstall = ReadString(iniFile, lang, "sInstall", sInstall);
            sRemove = ReadString(iniFile, lang, "sRemove", sRemove);
            sDone = ReadString(iniFile, lang, "sDone", sDone);
            sInput = ReadString(iniFile, lang, "sInput", sInput);
            sCount1Msg = ReadString(iniFile, lang, "sCount1Msg", sCount1Msg);
            sCount2Msg = ReadString(iniFile, lang, "sCount2Msg", sCount2Msg);
            sNoRoot = ReadString(iniFile, lang, "sNoRoot", sNoRoot);
            sAlreadyExists = ReadString(iniFile, lang, "sAlreadyExists", sAlreadyExists);
            sSetup = ReadString(iniFile, lang, "sSetup", sSetup);
            sImportInterface = ReadString(iniFile, lang, "sImportInterface", sImportInterface);
            sExportInterface = ReadString(iniFile, lang, "sExportInterface", sExportInterface);
            sQuickpickdialog = ReadString(iniFile, lang, "sQuickpickdialog", sQuickpickdialog);
            sQuicksavedialog = ReadString(iniFile, lang, "sQuicksavedialog", sQuicksavedialog);
            sFilesavedialog = ReadString(iniFile, lang, "sFilesavedialog", sFilesavedialog);
            sFileopendialog = ReadString(iniFile, lang, "sFileopendialog", sFileopendialog);
            sNoExportedViews = ReadString(iniFile, lang, "sNoExportedViews", sNoExportedViews);
            sIncludeSettings = ReadString(iniFile, lang, "sIncludeSettings", sIncludeSettings);
            sResetViewSetup = ReadString(iniFile, lang, "sResetViewSetup", sResetViewSetup);
            sResetViewNote = ReadString(iniFile, lang, "sResetViewNote", sResetViewNote);
        }

        static string ReadString(string iniFile, string section, string key, string defaultValue)
        {
            try
            {
                if (File.Exists(iniFile))
                {
                    return IniFileParser.ReadValue(section, key, defaultValue, iniFile);
                }
            }
            catch { }

            return defaultValue;
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
            string lang = ReadString(myIniFile, "General", "Lang", "");
            if (lang != "") return lang;

            lang = "en";

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
                        if (i < 3) { cmd += " \"%v\""; }
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

        static void DeleteAllValues(string NodeNum, string GUID)
        {
            string keyPath = $@"{ShellBagsKey}\{NodeNum}\Shell\{GUID}";

            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(keyPath, true))
                {
                    if (key != null)
                    {
                        string[] valueNames = key.GetValueNames();

                        foreach (string valueName in valueNames)
                        {
                            key.DeleteValue(valueName);
                        }
                    }
                }
            }
            catch (Exception)
            {
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


        // Get count of exported view reg files
        static int GetSavedViewsCount()
        {
            string[] regFiles = Directory.GetFiles(RegFolder, "*.reg");
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
                dialog.Icon = Icon.ExtractAssociatedIcon(myExe);
                dialog.StartPosition = FormStartPosition.Manual;
                dialog.AutoSize = true;
                dialog.AutoSizeMode = AutoSizeMode.GrowAndShrink;
                dialog.AutoScroll = true;


                if (Dark)
                {
                    DarkTitleBar(dialog.Handle);
                    dialog.BackColor = Color.FromArgb(32, 32, 32);
                    dialog.ForeColor = Color.White;
                }

                if (Directory.Exists(RegFolder))
                {
                    string[] regFiles = Directory.GetFiles(RegFolder, "*.reg");
                    if (regFiles.Length == 0)
                    {
                        CustomMessageBox.Show($"{sNoExportedViews}", sMain);
                        return null;
                    }

                    int yOffset = 0;
                    int maxWidth = 0;
                    int minWidth = (int)(120 * ScaleFactor);

                    foreach (string regFile in regFiles)
                    {
                        string itemName = Path.GetFileNameWithoutExtension(regFile);
                        int w = TextRenderer.MeasureText(itemName, new Font("Segoe UI", 10)).Width;
                        maxWidth = Math.Max(maxWidth, w);
                    }

                    if (maxWidth < minWidth) maxWidth = minWidth;

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
                Icon = Icon.ExtractAssociatedIcon(myExe);
                StartPosition = FormStartPosition.Manual;
                Width = (int)(420 * ScaleFactor);
                Height = (int)(230 * ScaleFactor);
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

                checkboxExp = new CheckBox();
                checkboxExp.Left = (int)(10 * ScaleFactor);
                checkboxExp.Top = textBoxInput.Bottom + (int)(10 * ScaleFactor);
                checkboxExp.Font = new Font("Segoe UI", 10);
                checkboxExp.Text = sIncludeSettings;
                checkboxExp.Checked = false;
                checkboxExp.AutoSize = true;
                Controls.Add(checkboxExp);

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
                    buttonOK.BackColor = Color.FromArgb(60, 60, 60);
                    buttonOK.FlatAppearance.MouseOverBackColor = Color.Black;
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

                Icon = Icon.ExtractAssociatedIcon(myExe);
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

                int h = TextRenderer.MeasureText(message, new Font("Segoe UI", 10)).Height;
                Height = Math.Max(Height, h);

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
                    buttonOK.BackColor = Color.FromArgb(60, 60, 60);
                    buttonOK.FlatAppearance.MouseOverBackColor = Color.Black;
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

                int w = TextRenderer.MeasureText(button2, new Font("Segoe UI", 10)).Height;
                b2Width = Math.Max(b2Width, w);

                message = $"\n{message}";

                Icon = Icon.ExtractAssociatedIcon(myExe);
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
                    buttonYes.BackColor = Color.FromArgb(60, 60, 60);
                    buttonYes.FlatAppearance.MouseOverBackColor = Color.Black;
                    buttonNo.FlatStyle = FlatStyle.Flat;
                    buttonNo.FlatAppearance.BorderColor = SystemColors.Highlight;
                    buttonNo.FlatAppearance.BorderSize = 1;
                    buttonNo.BackColor = Color.FromArgb(60, 60, 60);
                    buttonNo.FlatAppearance.MouseOverBackColor = Color.Black;
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
            private CheckBox checkboxRC;
            private Label labelRC;
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
                using (Font font = new Font("Segoe UI", 10))
                {
                    int setupWidth = TextRenderer.MeasureText(sResetViewSetup, font).Width;
                    int noteWidth = TextRenderer.MeasureText(sResetViewNote, font).Width;
                    int calculatedWidth = Math.Max(300, Math.Max(setupWidth, noteWidth)) + (int)(50 * ScaleFactor);
                    Width = calculatedWidth;
                }

                Text = sMain;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false;
                MinimizeBox = false;
                Icon = Icon.ExtractAssociatedIcon(myExe);
                StartPosition = FormStartPosition.Manual;
                Height = (int)(324 * ScaleFactor);
                Font = new Font("Segoe UI", 10);

                var importGroupBox = new GroupBox
                {
                    Name = "importGroupBox",
                    Text = sImportInterface,
                    Location = new Point((int)(10 * ScaleFactor), (int)(10 * ScaleFactor)),
                    Width = Width - (int)(40 * ScaleFactor),
                    Height = (int)(70 * ScaleFactor)
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
                    Height = (int)(70 * ScaleFactor)
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

                checkboxRC = new CheckBox();
                checkboxRC.Font = new Font("Segoe UI", 10);
                checkboxRC.Text = sResetViewSetup;
                checkboxRC.Checked = File.Exists(ResetViewFile);
                checkboxRC.AutoSize = true;
                checkboxRC.Left = (int)(10 * ScaleFactor);
                checkboxRC.Top = (int)(190 * ScaleFactor);
                checkboxRC.CheckedChanged += new EventHandler(CB1);

                labelRC = new Label();
                labelRC.Font = new Font("Segoe UI", 10);
                labelRC.Text = sResetViewNote;
                labelRC.AutoSize = true;
                labelRC.Left = (int)(26 * ScaleFactor);
                labelRC.Top = (int)(212 * ScaleFactor);

                Controls.Add(importGroupBox);
                Controls.Add(exportGroupBox);
                Controls.Add(checkboxRC);
                Controls.Add(labelRC);

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
                    buttonOK.BackColor = Color.FromArgb(60, 60, 60);
                    buttonOK.FlatAppearance.MouseOverBackColor = Color.Black;
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

            private void CB1(object sender, EventArgs e)
            {
                if (checkboxRC.Checked) { ExportExplorerSettings(ResetViewFile, true); } else { File.Delete(ResetViewFile); }
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

        // IMPORT function
        static void ImportRegistryFile(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return;
                }

                // Detect encoding based on the file content
                Encoding encoding = DetectFileEncoding(filePath);

                using (StreamReader reader = new StreamReader(filePath, encoding))
                {
                    string line;
                    RegistryKey currentKey = null;
                    string accumulatedLine = ""; // To accumulate lines with continuation characters

                    while ((line = reader.ReadLine()) != null)
                    {
                        line = line.Trim();

                        // Skip empty or comment lines
                        if (string.IsNullOrWhiteSpace(line) || line.StartsWith(";"))
                            continue;

                        // Handle line continuation (backslash at the end)
                        if (line.EndsWith("\\"))
                        {
                            accumulatedLine += line.Substring(0, line.Length - 1).Trim();
                            continue; // Wait for the next line to continue processing
                        }
                        else
                        {
                            // Accumulate the last part of the line
                            accumulatedLine += line.Trim();
                        }

                        // Now process the complete accumulated line
                        if (accumulatedLine.StartsWith("[") && accumulatedLine.EndsWith("]"))
                        {
                            // It's a key, so we process it as such
                            string keyPath = accumulatedLine.Trim('[', ']');
                            currentKey?.Close();

                            currentKey = CreateOrOpenRegistryKey(keyPath);
                        }
                        else if (currentKey != null)
                        {
                            ParseAndSetRegistryValue(currentKey, accumulatedLine);
                        }

                        accumulatedLine = "";
                    }

                    currentKey?.Close();
                }

                Console.WriteLine($"Successfully imported registry from file: {filePath}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error importing registry: {ex.Message}");
            }
        }

        // Function to detect encoding by reading the byte order mark (BOM) or file content
        static Encoding DetectFileEncoding(string filePath)
        {
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                if (file.Length >= 2)
                {
                    byte[] bom = new byte[2];
                    file.Read(bom, 0, 2);

                    // Check for UTF-16 LE BOM (FF FE)
                    if (bom[0] == 0xFF && bom[1] == 0xFE)
                        return Encoding.Unicode;

                    // Check for UTF-16 BE BOM (FE FF)
                    if (bom[0] == 0xFE && bom[1] == 0xFF)
                        return Encoding.BigEndianUnicode;
                }

                // Default to ANSI (UTF-8 without BOM)
                return Encoding.Default;
            }
        }

        static RegistryKey CreateOrOpenRegistryKey(string keyPath)
        {
            try
            {
                keyPath = NormalizeKeyPath(keyPath);

                int firstSlashIndex = keyPath.IndexOf('\\');
                if (firstSlashIndex == -1)
                {
                    // No subkey exists, just return the base key
                    return GetBaseRegistryKey(keyPath, writable: true, createSubKey: true);
                }

                string baseKeyName = keyPath.Substring(0, firstSlashIndex);
                string subKeyPath = keyPath.Substring(firstSlashIndex + 1);

                // Get the base registry key (e.g., HKEY_CURRENT_USER)
                RegistryKey baseKey = GetBaseRegistryKey(baseKeyName, writable: true, createSubKey: true);
                if (baseKey == null)
                {
                    throw new InvalidOperationException($"Invalid base registry key: {baseKeyName}");
                }

                // Create the subkey path if it doesn't exist
                return CreateSubKeyPath(baseKey, subKeyPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error creating or opening registry key '{keyPath}': {ex.Message}");
                return null;
            }
        }

        static RegistryKey CreateSubKeyPath(RegistryKey baseKey, string subKeyPath)
        {
            try
            {
                // Recursively create or open each subkey in the path
                RegistryKey currentKey = baseKey.CreateSubKey(subKeyPath, true);
                if (currentKey == null)
                {
                    throw new InvalidOperationException($"Failed to create or open registry subkey: {subKeyPath}");
                }

                return currentKey;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error creating subkey path '{subKeyPath}': {ex.Message}");
                return null;
            }
        }

        static void ParseAndSetRegistryValue(RegistryKey key, string line)
        {
            string name; string valueData; bool empty;

            void GetNameAndData(string pattern)
            {
                string[] parts = line.Split(new string[] { pattern }, 2, StringSplitOptions.None);
                name = parts[0].Substring(1).Replace("\\\"", "\"").Replace("\\\\", "\\");
                if (name == "@") name = null;
                valueData = parts[1].Substring(0, parts[1].Length);
                if (valueData.EndsWith("\"")) valueData = valueData.Substring(0, valueData.Length - 1);
                valueData = valueData.Replace("\\\"", "\"").Replace("\\\\", "\\");
                empty = valueData.Length == 0;
            }

            void TrimThis(string pattern)
            {
                if (valueData.EndsWith(pattern)) valueData = valueData.Substring(0, valueData.Length - pattern.Length);
                if (valueData == "00,00") valueData = "";
                empty = valueData.Length == 0;
            }

            if (line.StartsWith("@="))
            {
                line = $"\"@\"{line.Substring(1)}";
            }

            string regType = string.Empty;

            if (line.Contains("\"=\"")) regType = "REG_SZ";
            else if (line.Contains("\"=dword:")) regType = "REG_DWORD";
            else if (line.Contains("\"=hex(b):")) regType = "REG_QWORD";
            else if (line.Contains("\"=hex:")) regType = "REG_BINARY";
            else if (line.Contains("\"=hex(2):")) regType = "REG_EXPAND_SZ";
            else if (line.Contains("\"=hex(7):")) regType = "REG_MULTI_SZ";
            else if (line.Contains("\"=hex(0):")) regType = "REG_NONE";

            switch (regType)
            {
                case "REG_SZ":
                    GetNameAndData("\"=\"");
                    key.SetValue(name, valueData, RegistryValueKind.String);
                    break;

                case "REG_DWORD":
                    GetNameAndData("\"=dword:");
                    uint dwordValue = Convert.ToUInt32(valueData, 16);
                    key.SetValue(name, (int)dwordValue, RegistryValueKind.DWord);
                    break;

                case "REG_QWORD":
                    GetNameAndData("\"=hex(b):");
                    string[] bytePairs = valueData.Split(new[] { ',' });
                    ulong qwordValue = 0;
                    for (int i = 0; i < bytePairs.Length; i++)
                    {
                        qwordValue |= ((ulong)Convert.ToByte(bytePairs[i], 16)) << (i * 8);
                    }
                    key.SetValue(name, qwordValue, RegistryValueKind.QWord);
                    break;

                case "REG_BINARY":
                    GetNameAndData("\"=hex:");
                    if (empty)
                    {
                        key.SetValue(name, new byte[0], RegistryValueKind.Binary);
                    }
                    else
                    {
                        byte[] binaryData = valueData.Split(',').Select(h => Convert.ToByte(h, 16)).ToArray();
                        key.SetValue(name, binaryData, RegistryValueKind.Binary);
                    }
                    break;

                case "REG_EXPAND_SZ":
                    GetNameAndData("\"=hex(2):");
                    TrimThis(",00,00");
                    TrimThis(",00,00");

                    if (empty)
                    {
                        key.SetValue(name, string.Empty, RegistryValueKind.ExpandString);
                    }
                    else
                    {
                        byte[] binaryData = valueData.Split(',').Select(h => Convert.ToByte(h, 16)).ToArray();
                        string expandSzValue = Encoding.Unicode.GetString(binaryData);
                        key.SetValue(name, expandSzValue, RegistryValueKind.ExpandString);
                    }
                    break;

                case "REG_MULTI_SZ":
                    GetNameAndData("\"=hex(7):");
                    TrimThis(",00,00");
                    TrimThis(",00,00");

                    if (empty)
                    {
                        key.SetValue(name, new string[] { string.Empty }, RegistryValueKind.MultiString);
                    }
                    else
                    {
                        byte[] binaryData = valueData.Split(',').Select(h => Convert.ToByte(h, 16)).ToArray();
                        string multiSzValue = Encoding.Unicode.GetString(binaryData);
                        string[] multiStringArray = multiSzValue.Split(new[] { '\0' });
                        key.SetValue(name, multiStringArray, RegistryValueKind.MultiString);
                    }
                    break;

                case "REG_NONE":
                    GetNameAndData("\"=hex(0):");
                    if (empty)
                    {
                        key.SetValue(name, Encoding.Unicode.GetBytes("" + '\0'), RegistryValueKind.None);
                    }
                    else
                    {
                        byte[] binaryData = valueData.Split(',').Select(h => Convert.ToByte(h, 16)).ToArray();
                        key.SetValue(name, binaryData, RegistryValueKind.None);
                    }
                    break;

                default:
                    throw new InvalidOperationException("Unknown registry value type.");
            }
        }

        // EXPORT function
        static void ExportRegistryKey(string keyPath, string filePath, bool forceOverwrite)
        {
            keyPath = NormalizeKeyPath(keyPath);
            bool append = false;

            try
            {
                if (File.Exists(filePath) && !forceOverwrite) append = true;

                using (RegistryKey key = GetBaseRegistryKey(keyPath))
                {
                    if (key == null)
                    {
                        Console.Error.WriteLine($"Registry key not found: {keyPath}");
                        return;
                    }

                    using (StreamWriter writer = new StreamWriter(filePath, append, Encoding.Unicode))
                    {
                        if (!append) writer.WriteLine("Windows Registry Editor Version 5.00\r\n");
                        ExportKey(key, keyPath, writer);
                    }
                    Console.WriteLine($"Registry exported to file: {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error exporting registry: {ex.Message}");
            }
        }

        static string NormalizeKeyPath(string keyPath)
        {
            // map abbreviations and full names to their full uppercase versions
            var keyMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "HKLM", "HKEY_LOCAL_MACHINE" },
                { "HKEY_LOCAL_MACHINE", "HKEY_LOCAL_MACHINE" },
                { "HKCU", "HKEY_CURRENT_USER" },
                { "HKEY_CURRENT_USER", "HKEY_CURRENT_USER" },
                { "HKCR", "HKEY_CLASSES_ROOT" },
                { "HKEY_CLASSES_ROOT", "HKEY_CLASSES_ROOT" },
                { "HKU", "HKEY_USERS" },
                { "HKEY_USERS", "HKEY_USERS" },
                { "HKCC", "HKEY_CURRENT_CONFIG" },
                { "HKEY_CURRENT_CONFIG", "HKEY_CURRENT_CONFIG" }
            };

            string[] pathParts = keyPath.Split(new char[] { '\\' }, 2);

            if (keyMappings.ContainsKey(pathParts[0]))
            {
                pathParts[0] = keyMappings[pathParts[0]];
            }
            return pathParts.Length > 1 ? $"{pathParts[0]}\\{pathParts[1]}" : pathParts[0];
        }

        static void ExportKey(RegistryKey key, string keyPath, StreamWriter writer)
        {
            writer.WriteLine($"[{keyPath}]");
            foreach (string valueName in key.GetValueNames())
            {
                object value = key.GetValue(valueName, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                RegistryValueKind kind = key.GetValueKind(valueName);
                writer.WriteLine(FormatRegistryValue(valueName, value, kind));
            }
            writer.WriteLine();

            foreach (string subkeyName in key.GetSubKeyNames())
            {
                using (RegistryKey subKey = key.OpenSubKey(subkeyName))
                {
                    if (subKey != null)
                    {
                        ExportKey(subKey, $"{keyPath}\\{subkeyName}", writer);
                    }
                }
            }
        }

        static string FormatRegistryValue(string name, object value, RegistryValueKind kind)
        {
            string formattedName = string.IsNullOrEmpty(name) ? "@" : $"\"{name.Replace("\\", "\\\\").Replace("\"", "\\\"")}\"";

            switch (kind)
            {
                case RegistryValueKind.String:
                    return $"{formattedName}=\"{value.ToString().Replace("\\", "\\\\").Replace("\"", "\\\"")}\"";

                case RegistryValueKind.ExpandString:
                    string hexExpandString = string.Join(",", BitConverter.ToString(Encoding.Unicode.GetBytes(value.ToString())).Split('-'));
                    return $"{formattedName}=hex(2):{hexExpandString}";

                case RegistryValueKind.DWord:
                    return $"{formattedName}=dword:{((uint)(int)value).ToString("x8")}";

                case RegistryValueKind.QWord:
                    ulong qwordValue = Convert.ToUInt64(value);
                    return $"{formattedName}=hex(b):{string.Join(",", BitConverter.GetBytes(qwordValue).Select(b => b.ToString("x2")))}";

                case RegistryValueKind.Binary:
                    return $"{formattedName}=hex:{BitConverter.ToString((byte[])value).Replace("-", ",")}";

                case RegistryValueKind.MultiString:
                    string[] multiStrings = (string[])value;
                    var hexValues = multiStrings.SelectMany(s => Encoding.Unicode.GetBytes(s)
                                                    .Select(b => b.ToString("x2"))
                                                    .Concat(new[] { "00", "00" }))
                                                    .ToList();
                    hexValues.AddRange(new[] { "00", "00" });
                    return $"{formattedName}=hex(7):{string.Join(",", hexValues)}";

                case RegistryValueKind.None:
                    byte[] noneData = value as byte[];
                    string hexNone = string.Join(",", noneData.Select(b => b.ToString("x2")));
                    return $"{formattedName}=hex(0):{hexNone}";

                default:
                    throw new NotSupportedException($"Unsupported registry value type: {kind}");
            }
        }

        static RegistryKey GetBaseRegistryKey(string keyPath, bool writable = false, bool createSubKey = false)
        {
            string[] keyParts = keyPath.Split(new char[] { '\\' }, 2); // Split into root and subkey
            string root = keyParts[0].ToUpper();
            string subKey = keyParts.Length > 1 ? keyParts[1] : string.Empty; // If there's no subkey, handle it as empty

            RegistryKey baseKey;

            switch (root)
            {
                case "HKEY_LOCAL_MACHINE":
                    baseKey = Registry.LocalMachine;
                    break;
                case "HKEY_CURRENT_USER":
                    baseKey = Registry.CurrentUser;
                    break;
                case "HKEY_CLASSES_ROOT":
                    baseKey = Registry.ClassesRoot;
                    break;
                case "HKEY_USERS":
                    baseKey = Registry.Users;
                    break;
                case "HKEY_CURRENT_CONFIG":
                    baseKey = Registry.CurrentConfig;
                    break;
                default:
                    throw new NotSupportedException($"Unsupported root key: {root}");
            }

            if (string.IsNullOrEmpty(subKey)) return baseKey;

            if (createSubKey)
            {
                return baseKey.CreateSubKey(subKey, writable);
            }
            else
            {
                return baseKey.OpenSubKey(subKey, writable);
            }
        }

    }
}