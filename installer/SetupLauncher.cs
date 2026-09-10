using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;

[assembly: AssemblyTitle("Excel Navigator 安装程序")]
internal static class SetupLauncher
{
    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            bool extractOnly = args.Length == 2 && args[0] == "--extract-only";
            if (args.Length != 0 && !extractOnly) throw new ArgumentException("无效的安装参数。");
            bool acquired;
            using (var mutex = new Mutex(true, "Local\\ExcelNavigatorPane.Setup", out acquired))
            {
                if (!acquired) throw new InvalidOperationException("另一个 Excel Navigator 安装程序正在运行。");
                if (!extractOnly && MessageBox.Show(
                    "请先保存工作并关闭 Excel，然后继续安装。\n\n缺少 .NET Framework 4.8 或 VSTO Runtime 时需要联网下载，可能需要管理员权限。\n\n这是内部试用版，Windows 或 Office 可能显示发布者信任提示。",
                    "安装 Excel Navigator", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return 0;

                // Keep the ClickOnce source URL stable across repeated installs and upgrades.
                string root = Path.GetFullPath(extractOnly ? args[1] : Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelNavigatorPane", "Installer"));
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
                using (var setup = Process.Start(new ProcessStartInfo(Path.Combine(root, "setup.exe"))
                    { WorkingDirectory = root, UseShellExecute = true }))
                {
                    setup.WaitForExit();
                    return setup.ExitCode;
                }
            }
        }
        catch (Exception ex)
        {
            if (args.Length == 2 && args[0] == "--extract-only") Console.Error.WriteLine(ex.Message);
            else MessageBox.Show("安装未完成：" + ex.Message, "Excel Navigator", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 1;
        }
    }
}
