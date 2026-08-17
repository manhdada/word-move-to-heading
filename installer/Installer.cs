using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Word Move to Heading Setup")]
[assembly: AssemblyVersion("2.2.0.0")]
[assembly: AssemblyFileVersion("2.2.0.0")]

namespace WordMoveToHeadingSetup
{
    internal static class Program
    {
        private const string ProductName = "Word Move to Heading";
        private const string Version = "2.2.0";
        private const string ProgId = "WordMoveToHeading.Connect2";
        private const string ClassId = "{AD9CF34E-04AB-4628-B2E8-90CA487BC348}";
        private const string OldProgId = "WordMoveToHeading.Connect";
        private const string OldClassId = "{7D4B0E14-46AB-4F27-BB25-1391E1B42413}";
        private const string AssemblyName = "WordMoveToHeading, Version=2.2.0.0, Culture=neutral, PublicKeyToken=null";
        private const string ClassName = "WordMoveToHeading.Connect";

        [STAThread]
        private static int Main(string[] args)
        {
            bool uninstall = HasArgument(args, "/uninstall");
            bool silent = HasArgument(args, "/silent");

            try
            {
                if (Process.GetProcessesByName("WINWORD").Length > 0)
                {
                    if (!silent)
                    {
                        MessageBox.Show(
                            "Hãy đóng hoàn toàn Microsoft Word rồi chạy lại bộ cài.",
                            ProductName,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    return 2;
                }

                if (uninstall)
                {
                    Uninstall();
                    if (!silent)
                    {
                        MessageBox.Show("Đã gỡ " + ProductName + ".", ProductName,
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    ScheduleSelfRemoval();
                    return 0;
                }

                Install();
                if (!silent)
                {
                    MessageBox.Show(
                        "Cài đặt phiên bản " + Version + " thành công.\r\n\r\n" +
                        "Mở Word để dùng Move to và nút Nhận diện Heading trên tab Home.",
                        ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                return 0;
            }
            catch (Exception exception)
            {
                if (!silent)
                {
                    MessageBox.Show(
                        "Cài đặt không thành công:\r\n" + exception.Message,
                        ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                return 1;
            }
        }

        private static void Install()
        {
            string targetDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WordMoveToHeading");
            string targetDll = Path.Combine(targetDirectory, "WordMoveToHeading.dll");
            string uninstaller = Path.Combine(targetDirectory, "Uninstall.exe");

            Directory.CreateDirectory(targetDirectory);
            RemoveRegistration();

            using (Stream input = Assembly.GetExecutingAssembly().GetManifestResourceStream("WordMoveToHeading.Payload.dll"))
            {
                if (input == null)
                {
                    throw new InvalidOperationException("Không tìm thấy DLL add-in trong bộ cài.");
                }
                using (FileStream output = new FileStream(targetDll, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    input.CopyTo(output);
                }
            }

            File.Copy(Application.ExecutablePath, uninstaller, true);
            string codeBase = new Uri(targetDll).AbsoluteUri;

            RegisterView(RegistryView.Registry32, codeBase);
            RegisterView(RegistryView.Registry64, codeBase);

            using (RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
            using (RegistryKey key = baseKey.CreateSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Uninstall\WordMoveToHeading"))
            {
                key.SetValue("DisplayName", ProductName, RegistryValueKind.String);
                key.SetValue("DisplayVersion", Version, RegistryValueKind.String);
                key.SetValue("Publisher", "Local installation", RegistryValueKind.String);
                key.SetValue("InstallLocation", targetDirectory, RegistryValueKind.String);
                key.SetValue("UninstallString", "\"" + uninstaller + "\" /uninstall", RegistryValueKind.String);
                key.SetValue("NoModify", 1, RegistryValueKind.DWord);
                key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
            }
        }

        private static void RegisterView(RegistryView view, string codeBase)
        {
            using (RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
            {
                string inprocPath = @"Software\Classes\CLSID\" + ClassId + @"\InprocServer32";
                using (RegistryKey key = baseKey.CreateSubKey(inprocPath))
                {
                    SetComValues(key, codeBase);
                }
                using (RegistryKey key = baseKey.CreateSubKey(inprocPath + @"\2.2.0.0"))
                {
                    SetComValues(key, codeBase);
                }
                using (RegistryKey key = baseKey.CreateSubKey(@"Software\Classes\" + ProgId + @"\CLSID"))
                {
                    key.SetValue(null, ClassId, RegistryValueKind.String);
                }
                using (RegistryKey key = baseKey.CreateSubKey(
                    @"Software\Microsoft\Office\Word\Addins\" + ProgId))
                {
                    key.SetValue("FriendlyName", ProductName, RegistryValueKind.String);
                    key.SetValue("Description", "Move text to headings and auto-detect outline levels.", RegistryValueKind.String);
                    key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    key.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
                }
                using (RegistryKey key = baseKey.CreateSubKey(
                    @"Software\Microsoft\Office\16.0\Word\Resiliency\DoNotDisableAddinList"))
                {
                    key.SetValue(ProgId, 1, RegistryValueKind.DWord);
                }
            }
        }

        private static void SetComValues(RegistryKey key, string codeBase)
        {
            key.SetValue(null, "mscoree.dll", RegistryValueKind.String);
            key.SetValue("ThreadingModel", "Both", RegistryValueKind.String);
            key.SetValue("Class", ClassName, RegistryValueKind.String);
            key.SetValue("Assembly", AssemblyName, RegistryValueKind.String);
            key.SetValue("RuntimeVersion", "v4.0.30319", RegistryValueKind.String);
            key.SetValue("CodeBase", codeBase, RegistryValueKind.String);
        }

        private static void Uninstall()
        {
            RemoveRegistration();
            string targetDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "WordMoveToHeading");
            string dll = Path.Combine(targetDirectory, "WordMoveToHeading.dll");
            if (File.Exists(dll))
            {
                File.Delete(dll);
            }
        }

        private static void RemoveRegistration()
        {
            foreach (RegistryView view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
            {
                using (RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
                {
                    DeleteTree(baseKey, @"Software\Classes\CLSID\" + ClassId);
                    DeleteTree(baseKey, @"Software\Classes\CLSID\" + OldClassId);
                    DeleteTree(baseKey, @"Software\Classes\" + ProgId);
                    DeleteTree(baseKey, @"Software\Classes\" + OldProgId);
                    DeleteTree(baseKey, @"Software\Microsoft\Office\Word\Addins\" + ProgId);
                    DeleteTree(baseKey, @"Software\Microsoft\Office\Word\Addins\" + OldProgId);
                    using (RegistryKey resilient = baseKey.OpenSubKey(
                        @"Software\Microsoft\Office\16.0\Word\Resiliency\DoNotDisableAddinList", true))
                    {
                        if (resilient != null)
                        {
                            resilient.DeleteValue(ProgId, false);
                            resilient.DeleteValue(OldProgId, false);
                        }
                    }
                }
            }

            using (RegistryKey baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
            {
                DeleteTree(baseKey, @"Software\Microsoft\Windows\CurrentVersion\Uninstall\WordMoveToHeading");
            }
        }

        private static void DeleteTree(RegistryKey baseKey, string path)
        {
            try { baseKey.DeleteSubKeyTree(path, false); }
            catch (ArgumentException) { }
        }

        private static bool HasArgument(string[] args, string expected)
        {
            foreach (string arg in args)
            {
                if (string.Equals(arg, expected, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        private static void ScheduleSelfRemoval()
        {
            string targetDirectory = Path.GetDirectoryName(Application.ExecutablePath);
            if (string.IsNullOrEmpty(targetDirectory))
            {
                return;
            }

            var startInfo = new ProcessStartInfo(
                "cmd.exe",
                "/c ping 127.0.0.1 -n 3 > nul & rmdir /s /q \"" + targetDirectory + "\"");
            startInfo.CreateNoWindow = true;
            startInfo.UseShellExecute = false;
            Process.Start(startInfo);
        }
    }
}

