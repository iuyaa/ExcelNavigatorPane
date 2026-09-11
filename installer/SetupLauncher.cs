using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Excel Navigator 安装程序")]
internal static class SetupLauncher
{
    private static string diagnosticFile;
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            bool extractOnly = args.Length == 2 && args[0] == "--extract-only";
            bool checkOnly = args.Length == 1 && args[0] == "--check-only";
            if (args.Length != 0 && !extractOnly && !checkOnly) throw new ArgumentException("无效的安装参数。");
            if (!extractOnly)
            {
                string logs = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelNavigatorPane", "Installer", "logs");
                Directory.CreateDirectory(logs);
                diagnosticFile = Path.Combine(logs, "setup-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fff") + "-" + Process.GetCurrentProcess().Id + ".log");
                File.AppendAllText(diagnosticFile, "Version: " + Assembly.GetExecutingAssembly().GetName().Version + "\r\nExecutable: " + Assembly.GetExecutingAssembly().Location + "\r\nUser: " + System.Security.Principal.WindowsIdentity.GetCurrent().Name + "\r\n");
            }
            if (checkOnly) { EnsureNoLegacyRegistration(); return 0; }
            bool acquired;
            using (var mutex = new Mutex(true, "Local\\ExcelNavigatorPane.Setup", out acquired))
            {
                if (!acquired) throw new InvalidOperationException("另一个 Excel Navigator 安装程序正在运行。");
                string architecture = null;
                if (!extractOnly)
                {
                    EnsureOfficeClosed();
                    DetectExcelArchitecture(); // Validate the supported host; payload is AnyCPU.
                    architecture = SelectInstallerArchitecture(Environment.Is64BitOperatingSystem);
                }
                if (!extractOnly && MessageBox.Show(
                    "将安装到 Program Files，需要管理员授权。\n\n缺少 .NET Framework 4.8 或 VSTO Runtime 时会从微软下载。\n\n安装过程中请保持 Excel 和 WPS 表格关闭；安装器不会导入根证书或自动关闭这些程序。",
                    "安装 Excel Navigator", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return 0;
                if (!extractOnly) EnsureNoLegacyRegistration(offerDevelopmentMigration: true);

                // Retain versioned MSI sources for Windows Installer repair.
                string root = Path.GetFullPath(extractOnly ? args[1] : Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelNavigatorPane", "Installer",
                    Assembly.GetExecutingAssembly().GetName().Version.ToString()));
                if (extractOnly && Directory.Exists(root)) throw new IOException("检查目录必须尚不存在，以免覆盖文件。");
                Directory.CreateDirectory(root);
                using (var resource = Assembly.GetExecutingAssembly().GetManifestResourceStream("payload.zip"))
                using (var archive = new ZipArchive(resource, ZipArchiveMode.Read))
                {
                    foreach (var entry in archive.Entries)
                    {
                        string destination = Path.GetFullPath(Path.Combine(root, entry.FullName));
                        if (!destination.StartsWith(root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                            throw new InvalidDataException("安装文件路径无效。");
                        Directory.CreateDirectory(Path.GetDirectoryName(destination));
                        if (entry.Name.Length == 0) continue;
                        using (var input = entry.Open())
                        using (var output = File.Create(destination)) input.CopyTo(output);
                    }
                }
                if (extractOnly) return 0;
                EnsureOfficeClosed();
                string folder = Path.Combine(root, architecture);
                using (var setup = Process.Start(new ProcessStartInfo(Path.Combine(folder, "setup.exe"))
                    { WorkingDirectory = folder, UseShellExecute = true }))
                {
                    setup.WaitForExit();
                    if (setup.ExitCode != 0) return setup.ExitCode;
                }
                EnsureOfficeClosed();
                string msi = Path.Combine(folder, "ExcelNavigator-" + Assembly.GetExecutingAssembly().GetName().Version + "-" + architecture + ".msi");
                using (var install = Process.Start(new ProcessStartInfo(Path.Combine(Environment.SystemDirectory, "msiexec.exe"),
                    "/i \"" + msi + "\" /passive /norestart") { WorkingDirectory = folder, UseShellExecute = true, Verb = "runas" }))
                {
                    install.WaitForExit();
                    string message = InstallationResultMessage(install.ExitCode);
                    if (message != null) MessageBox.Show(message, "Excel Navigator", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return install.ExitCode;
                }
            }
        }
        catch (OperationCanceledException) { return 0; }
        catch (Exception ex)
        {
            var nativeError = ex as System.ComponentModel.Win32Exception;
            if (nativeError != null && nativeError.NativeErrorCode == 1223) return 0;
            if (diagnosticFile != null) { try { File.AppendAllText(diagnosticFile, ex + "\r\n"); } catch { } }
            if ((args.Length == 2 && args[0] == "--extract-only") || (args.Length == 1 && args[0] == "--check-only")) Console.Error.WriteLine(ex.Message);
            else MessageBox.Show("安装未完成：" + ex.Message + (diagnosticFile == null ? "" : "\n\n诊断日志：" + diagnosticFile), "Excel Navigator " + Assembly.GetExecutingAssembly().GetName().Version, MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }

    internal static void EnsureOfficeClosed()
    {
        WaitForOfficeClosed(IsOfficeRunning, () => MessageBox.Show(
            "请先保存所有工作，再关闭全部 Excel 和 WPS 表格窗口。\n\n关闭后点击“重试”继续安装；点击“取消”退出安装。",
            "Excel Navigator", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning));
    }

    internal static void WaitForOfficeClosed(Func<bool> isRunning, Func<DialogResult> prompt)
    {
        while (isRunning())
            if (prompt() != DialogResult.Retry) throw new OperationCanceledException();
    }

    internal static string InstallationResultMessage(int exitCode)
    {
        if (exitCode == 1602) return null;
        if (exitCode == 0) return "Excel Navigator 安装成功。\n\n现在可以重新打开 Excel 使用导航栏。";
        if (exitCode == 3010 || exitCode == 1641)
            return "Excel Navigator 安装成功。\n\nWindows 提示需要重启，请保存其他工作，重启电脑后再打开 Excel。";
        throw new InvalidOperationException("安装失败，Windows Installer 返回代码：" + exitCode + "。");
    }

    private static bool IsOfficeRunning()
    {
        bool running = false;
        foreach (string name in new[] { "EXCEL", "et" })
        {
            var processes = Process.GetProcessesByName(name);
            running |= processes.Length != 0;
            foreach (var process in processes) process.Dispose();
        }
        return running;
    }

    internal static string ReadPeArchitecture(string path)
    {
        using (var reader = new BinaryReader(File.OpenRead(path)))
        {
            if (reader.ReadUInt16() != 0x5a4d) throw new InvalidDataException("Excel 程序格式不正确。");
            reader.BaseStream.Position = 0x3c;
            int offset = reader.ReadInt32();
            if (offset < 0x40 || offset > reader.BaseStream.Length - 6) throw new InvalidDataException("Excel 程序头无效。");
            reader.BaseStream.Position = offset;
            if (reader.ReadUInt32() != 0x00004550) throw new InvalidDataException("Excel 程序头无效。");
            ushort machine = reader.ReadUInt16();
            if (machine == 0x8664) return "x64";
            if (machine == 0x014c) return "x86";
            throw new InvalidOperationException("本安装包暂不支持该 Excel 架构。");
        }
    }

    internal static string SelectInstallerArchitecture(bool is64BitWindows)
    {
        return is64BitWindows ? "x64" : "x86";
    }

    private static string DetectExcelArchitecture()
    {
        foreach (var hive in new[] { RegistryHive.LocalMachine, RegistryHive.CurrentUser })
        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            if (view == RegistryView.Registry64 && !Environment.Is64BitOperatingSystem) continue;
            using (var root = RegistryKey.OpenBaseKey(hive, view))
            using (var key = root.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"))
            {
                string path = Convert.ToString(key == null ? null : key.GetValue(null)).Trim('"');
                if (File.Exists(path)) return ReadPeArchitecture(path);
            }
        }
        throw new InvalidOperationException("未找到桌面版 Excel。请先安装 Excel，或使用与 Excel 位数一致的 MSI 安装包。");
    }

    private static void EnsureNoLegacyRegistration(bool offerDevelopmentMigration = false)
    {
        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            if (view == RegistryView.Registry64 && !Environment.Is64BitOperatingSystem) continue;
            using (var root = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view))
            using (var key = root.OpenSubKey(@"Software\Microsoft\Office\Excel\Addins\ExcelNavigatorPane"))
            using (var machine = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
            using (var installed = machine.OpenSubKey(@"Software\Microsoft\Office\Excel\Addins\ExcelNavigatorPane"))
            {
                string userManifest = Convert.ToString(key == null ? null : key.GetValue("Manifest"));
                string machineManifest = Convert.ToString(installed == null ? null : installed.GetValue("Manifest"));
                if (diagnosticFile != null) File.AppendAllText(diagnosticFile, view + " HKCU Manifest: " + userManifest + "\r\n" + view + " HKLM Manifest: " + machineManifest + "\r\n");
                if (HasConflictingRegistration(userManifest, machineManifest))
                {
                    if (offerDevelopmentMigration && IsDevelopmentRegistration(userManifest) && MessageBox.Show(
                        "发现开发版加载项注册：\n" + userManifest + "\n\n是否备份并移除此开发版注册，继续安装？\n只更改此加载项的用户注册，不会删除项目、工作簿或插件文件。\n备份保留在 HKCU\\Software\\ExcelNavigatorPane\\RegistrationBackups。",
                        "迁移开发版注册", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        EnsureOfficeClosed();
                        BackupAndRemoveDevelopmentRegistration(root, @"Software\Microsoft\Office\Excel\Addins\ExcelNavigatorPane", userManifest);
                        continue;
                    }
                    throw new InvalidOperationException("用户级加载项指向另一份安装，需先处理该注册：\nHKCU\\Software\\Microsoft\\Office\\Excel\\Addins\\ExcelNavigatorPane（" + view + "）\nManifest：" + userManifest + "\n不会自动删除注册或文件。请将此信息提供给维护人员。");
                }
            }
        }
    }

    internal static bool IsDevelopmentRegistration(string manifest)
    {
        const string suffix = "|vstolocal";
        Uri uri;
        if (string.IsNullOrEmpty(manifest) || !manifest.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) ||
            !Uri.TryCreate(manifest.Substring(0, manifest.Length - suffix.Length), UriKind.Absolute, out uri) ||
            !uri.IsFile || uri.IsUnc || uri.Query.Length != 0 || uri.Fragment.Length != 0) return false;
        string path = uri.LocalPath;
        if (!string.Equals(Path.GetFileName(path), "ExcelNavigatorPane.vsto", StringComparison.OrdinalIgnoreCase)) return false;
        var configuration = new DirectoryInfo(Path.GetDirectoryName(path));
        return (configuration.Name.Equals("Debug", StringComparison.OrdinalIgnoreCase) || configuration.Name.Equals("Release", StringComparison.OrdinalIgnoreCase)) &&
            configuration.Parent != null && configuration.Parent.Name.Equals("bin", StringComparison.OrdinalIgnoreCase);
    }

    internal static void BackupAndRemoveDevelopmentRegistration(RegistryKey root, string path, string expectedManifest)
    {
        if (!IsDevelopmentRegistration(expectedManifest)) throw new InvalidOperationException("仅支持迁移已识别的开发版注册。");
        using (var current = root.OpenSubKey(path))
        {
            if (current == null) return;
            if (!string.Equals(Convert.ToString(current.GetValue("Manifest")), expectedManifest, StringComparison.Ordinal))
                throw new InvalidOperationException("注册已变化，未移除，请重新检查。");
            string backupPath = @"Software\ExcelNavigatorPane\RegistrationBackups\" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N");
            using (var backup = root.CreateSubKey(backupPath))
            {
                CopyRegistration(current, backup);
                backup.Flush();
            }
            if (diagnosticFile != null) File.AppendAllText(diagnosticFile, "Development registration backup: HKCU\\" + backupPath + "\r\n");
            if (!string.Equals(Convert.ToString(current.GetValue("Manifest")), expectedManifest, StringComparison.Ordinal))
                throw new InvalidOperationException("备份期间注册已变化，保留原注册，请重新检查。");
        }
        root.DeleteSubKeyTree(path, false);
    }

    private static void CopyRegistration(RegistryKey source, RegistryKey target)
    {
        foreach (string name in source.GetValueNames())
            target.SetValue(name, source.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames), source.GetValueKind(name));
        foreach (string name in source.GetSubKeyNames())
            using (var child = source.OpenSubKey(name))
            using (var copy = target.CreateSubKey(name)) CopyRegistration(child, copy);
    }

    internal static bool HasConflictingRegistration(string userManifest, string machineManifest)
    {
        if (string.IsNullOrWhiteSpace(userManifest)) return false;
        // Equal local paths are the same installation, not a competing ClickOnce/development copy.
        const string suffix = "|vstolocal";
        Uri user, machine;
        if (!string.IsNullOrEmpty(machineManifest) && userManifest.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
            machineManifest.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) &&
            Uri.TryCreate(userManifest.Substring(0, userManifest.Length - suffix.Length), UriKind.Absolute, out user) &&
            Uri.TryCreate(machineManifest.Substring(0, machineManifest.Length - suffix.Length), UriKind.Absolute, out machine) &&
            user.IsFile && machine.IsFile && !user.IsUnc && !machine.IsUnc &&
            string.Equals(user.LocalPath, machine.LocalPath, StringComparison.OrdinalIgnoreCase)) return false;
        return true;
    }
}
