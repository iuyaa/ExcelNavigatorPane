using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Security.Principal;
using System.Linq;

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
            string operation = args.Length == 0 ? "install" : args.Length == 1 ? args[0].Replace("--apply-", "").Replace("--", "") : "";
            bool elevated = args.Length == 1 && args[0].StartsWith("--apply-", StringComparison.Ordinal);
            if (!extractOnly && !checkOnly && operation != "install" && operation != "restore" && operation != "uninstall")
                throw new ArgumentException("无效的安装参数。");
            if (extractOnly) { ExtractPayload(args[1],true); return 0; }
            string logs=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"ExcelNavigatorPane","Installer","logs");
            Directory.CreateDirectory(logs);
            diagnosticFile=Path.Combine(logs,"setup-"+DateTime.UtcNow.ToString("yyyyMMdd-HHmmss-fff")+"-"+Process.GetCurrentProcess().Id+".log");
            File.AppendAllText(diagnosticFile,"Version: "+Assembly.GetExecutingAssembly().GetName().Version+"\r\nOperation: "+operation+"\r\n");
            if(checkOnly){EnsureNoLegacyRegistration();return 0;}
            string root=InstallationRoot();
            string publicKey;
            using(var reader=new StreamReader(Assembly.GetExecutingAssembly().GetManifestResourceStream("update-public-key.xml"))) publicKey=reader.ReadToEnd();
            var scopes=OpenMachineScopes();
            try
            {
                var installation=new VersionedInstallation(root,publicKey,scopes);
                if(!elevated)
                {
                    if(operation=="uninstall")EnsureOfficeClosed("卸载需要释放插件文件。");
                    if(operation=="install" && installation.NeedsMigration())EnsureOfficeClosed("本次首次迁移旧安装，需要关闭一次；迁移后的常规更新无需关闭。");
                    else if(operation=="install" && !HavePrerequisites())EnsureOfficeClosed("本次需要安装运行环境，暂不能保留运行中的宿主。");
                    PrepareUserRegistration(installation,false);
                    if(MessageBox.Show(operation=="install"
                        ? "将安装到 Program Files，需要管理员授权。\n\n首次迁移旧版或安装运行环境时需要关闭 Excel/WPS；后续版本可继续工作，新版在新的 Excel/WPS 进程启动时生效。\n\n不会替你保存文件或关闭窗口。"
                        : operation=="restore" ? "将恢复上一版，当前窗口继续使用已加载版本，重新启动 Excel/WPS 后生效。需要管理员授权。"
                        : "将卸载 Excel Navigator 及保留的版本，需要管理员授权。", "Excel Navigator",MessageBoxButtons.OKCancel,MessageBoxIcon.Information)!=DialogResult.OK)return 0;
                    int result=StartElevated(operation);
                    if(result==0)PrepareUserRegistration(installation,true);
                    return result;
                }
                if(!new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator))
                    throw new UnauthorizedAccessException("维护操作需要管理员授权。");
                bool acquired;
                using(var mutex=new Mutex(true,"Global\\ExcelNavigatorPane.Setup",out acquired))
                {
                    if(!acquired)throw new InvalidOperationException("另一个 Excel Navigator 安装或维护程序正在运行。");
                    installation.RecoverActivation();
                    using(var machine=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,Environment.Is64BitOperatingSystem?RegistryView.Registry64:RegistryView.Registry32))
                    using(var session=machine.OpenSubKey(@"SYSTEM\CurrentControlSet\Control\Session Manager"))
                        if(PendingPathsAffect(root,session==null?null:session.GetValue("PendingFileRenameOperations") as string[]))
                            throw new InvalidOperationException("上次安装或卸载仍有文件等待重启处理，请先重启电脑再继续。");
                    installation.RecoverMigration(VersionedInstallation.Products("{A14219F5-2541-441C-BAEC-606015D32B55}").Count!=0);
                    if(operation=="uninstall")
                    {
                        EnsureOfficeClosed();
                        installation.EnsureOwnedRegistration();
                        // Keep registration until every payload is removed; failed uninstall remains retryable.
                        string active=installation.ReadState("CurrentVersion");
                        if(active.Length==0)throw new InvalidOperationException("没有新版安装记录，请通过原有卸载入口处理旧版。");
                        bool reboot=false;
                        foreach(string product in VersionedInstallation.Products(VersionedInstallation.PayloadFamily)
                            .OrderBy(p=>VersionedInstallation.ProductVersion(p)==(active.Length==0?"":Version.Parse(active).ToString(3))?1:0))
                        {
                            int result=RunMsi("/x "+product);RequireMsiSuccess(result,"卸载");reboot|=result!=0;
                        }
                        installation.RemoveRegistration();
                        CleanupMaintenance(root);
                        MessageBox.Show(reboot?"Excel Navigator 已卸载，Windows 要求重启以完成文件清理，请重启后再安装。":"Excel Navigator 已卸载。当前运行的维护程序缓存可在退出后手动删除。","Excel Navigator");return 0;
                    }
                    if(operation=="restore")
                    {
                        string previous=installation.ReadState("PreviousVersion");
                        if(previous.Length==0)throw new InvalidOperationException("没有可恢复的上一版。");
                        installation.Activate(previous,true);
                        MessageBox.Show("已恢复到 v"+previous+"。当前窗口保持原样，新的 Excel/WPS 进程将加载恢复的版本。","Excel Navigator");return 0;
                    }
                    DetectExcelArchitecture();
                    if(installation.NeedsMigration() || VersionedInstallation.Products("{A14219F5-2541-441C-BAEC-606015D32B55}").Count!=0)
                        EnsureOfficeClosed("本次首次迁移旧安装，需要关闭一次；迁移后的常规更新无需关闭。");
                    string legacyBackup=installation.BackupLegacy(Assembly.GetExecutingAssembly().Location);
                    string recoveryVersion=legacyBackup??installation.ReadState("CurrentVersion");
                    if(legacyBackup!=null)installation.StageMigration(legacyBackup);
                    string version=Assembly.GetExecutingAssembly().GetName().Version.ToString();
                    bool repair=false;
                    foreach(string product in VersionedInstallation.Products(VersionedInstallation.PayloadFamily))
                    {
                        Version installed=Version.Parse(VersionedInstallation.ProductVersion(product));
                        if(installed>Version.Parse(version))throw new InvalidOperationException("已安装更高版本；如需恢复请使用“恢复上一版”。");
                        repair|=installed.ToString(3)==Version.Parse(version).ToString(3);
                    }
                    if(repair)EnsureOfficeClosed("本次重新安装同一版本，需要修复已有文件；安装更高版本无需关闭。");
                    string versions=Path.Combine(root,"Versions");
                    if(Directory.Exists(versions))VersionedInstallation.RejectReparse(versions);
                    string target=installation.VersionFolder(version);
                    if(Directory.Exists(target))
                    {
                        VersionedInstallation.RejectReparse(target);
                        foreach(string file in Directory.GetFiles(target))VersionedInstallation.RejectReparse(file);
                    }
                    // The machine directory is private to elevated setup, unlike a user's extraction cache.
                    string cache=Path.Combine(root,"Installer",version);
                    if(Directory.Exists(cache))VersionedInstallation.RejectReparse(cache);
                    ExtractPayload(cache,false);
                    string architecture=SelectInstallerArchitecture(Environment.Is64BitOperatingSystem);
                    string folder=Path.Combine(cache,architecture);
                    if(!HavePrerequisites())
                    {
                        EnsureOfficeClosed();
                        using(var setup=Process.Start(new ProcessStartInfo(Path.Combine(folder,"setup.exe")){WorkingDirectory=folder,UseShellExecute=true}))
                        {
                            setup.WaitForExit();RequireMsiSuccess(setup.ExitCode,"安装运行环境");
                            if(setup.ExitCode!=0)throw new InvalidOperationException("运行环境要求重启，请重启电脑后重新运行安装程序。加载入口尚未切换。");
                        }
                        if(!HavePrerequisites())throw new InvalidOperationException("运行环境尚未安装完成，请完成安装或重启后重试。");
                    }
                    installation.StoreManager(version,Assembly.GetExecutingAssembly().Location);
                    string msi=Path.Combine(folder,"ExcelNavigator-"+version+"-"+architecture+".msi");
                    int code=RunMsi("/i \""+msi+"\""+(repair?" REINSTALL=ALL REINSTALLMODE=vomus":""));
                    if(code==1602)return 0;
                    try
                    {
                        RequireMsiSuccess(code,"安装");
                        if(code!=0)throw new InvalidOperationException("Windows 要求重启以完成文件安装。请重启后重新运行安装程序。");
                        installation.Activate(version,false,legacyBackup);
                        installation.ClearMigration();
                    }
                    catch
                    {
                        // Legacy removal is an MSI transaction; selector failure after its commit needs the signed backup.
                        if(!string.IsNullOrEmpty(recoveryVersion) && !installation.HasActiveRegistration(recoveryVersion) &&
                            VersionedInstallation.Products("{A14219F5-2541-441C-BAEC-606015D32B55}").Count==0)
                            installation.Activate(recoveryVersion,true);
                        throw;
                    }
                    // shortcut: keep all old payloads until uninstall; bounded cleanup needs cross-user usage detection.
                    MessageBox.Show(InstallationResultMessage(0),"Excel Navigator",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    return 0;
                }
            }
            finally{foreach(var scope in scopes)scope.Dispose();}
        }
        catch(OperationCanceledException){return 0;}
        catch(Exception ex)
        {
            var native=ex as System.ComponentModel.Win32Exception;
            if(native!=null&&native.NativeErrorCode==1223)return 0;
            if(diagnosticFile!=null){try{File.AppendAllText(diagnosticFile,ex+"\r\n");}catch{}}
            if(args.Length>0&&(args[0]=="--extract-only"||args[0]=="--check-only"))Console.Error.WriteLine(ex.Message);
            else MessageBox.Show("操作未完成："+ex.Message+(diagnosticFile==null?"":"\n\n诊断日志："+diagnosticFile),"Excel Navigator "+Assembly.GetExecutingAssembly().GetName().Version,MessageBoxButtons.OK,MessageBoxIcon.Error);
            return 1;
        }
    }

    internal static string InstallationRoot()
    {
        using(var machine=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,Environment.Is64BitOperatingSystem?RegistryView.Registry64:RegistryView.Registry32))
        using(var windows=machine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion"))
        {
            string programFiles=Convert.ToString(windows.GetValue("ProgramFilesDir"));
            if(!Path.IsPathRooted(programFiles)||!Directory.Exists(programFiles))throw new InvalidDataException("Program Files 路径无效。");
            return Path.Combine(programFiles,"ExcelNavigatorPane");
        }
    }
    static int StartElevated(string operation)
    {
        // Mapped/subst drives can disappear under UAC; run the identical launcher from a local path.
        string cache=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"ExcelNavigatorPane","Installer","Launchers",Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(cache);VersionedInstallation.RejectReparse(cache);
        string launcher=Path.Combine(cache,"ExcelNavigator-Setup.exe");
        try
        {
            File.Copy(Assembly.GetExecutingAssembly().Location,launcher,false);
            using(var held=new FileStream(launcher,FileMode.Open,FileAccess.Read,FileShare.Read))
            using(var child=Process.Start(new ProcessStartInfo(launcher,"--apply-"+operation){UseShellExecute=true,Verb="runas"}))
            {child.WaitForExit();return child.ExitCode;}
        }
        finally
        {
            try{if(File.Exists(launcher))File.Delete(launcher);if(Directory.GetFileSystemEntries(cache).Length==0)Directory.Delete(cache);}
            catch(IOException){ }
        }
    }
    internal static bool PendingPathsAffect(string root,string[] paths)
    {
        string prefix=Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar)+Path.DirectorySeparatorChar;
        return paths!=null&&paths.Any(path=>!string.IsNullOrEmpty(path)&&
            (path.StartsWith(@"\??\",StringComparison.Ordinal)?path.Substring(4):path).StartsWith(prefix,StringComparison.OrdinalIgnoreCase));
    }
    static RegistryKey[] OpenMachineScopes()
    {
        return Environment.Is64BitOperatingSystem
            ? new[]{RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry64),RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry32)}
            : new[]{RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry32)};
    }
    static bool HavePrerequisites()
    {
        using(var machine=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,RegistryView.Registry32))
        using(var net=machine.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
            if(net==null || Convert.ToInt32(net.GetValue("Release",0))<528040)return false;
        foreach(var view in new[]{RegistryView.Registry32,RegistryView.Registry64})
        {
            if(view==RegistryView.Registry64&&!Environment.Is64BitOperatingSystem)continue;
            using(var machine=RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,view))
            foreach(string path in new[]{@"SOFTWARE\Microsoft\VSTO Runtime Setup\v4R",@"SOFTWARE\Microsoft\VSTO Runtime Setup\v4"})
                using(var runtime=machine.OpenSubKey(path))if(runtime!=null&&!string.IsNullOrEmpty(Convert.ToString(runtime.GetValue("Version"))))return true;
        }
        return false;
    }
    static void ExtractPayload(string root,bool extractOnly)
    {
        root=Path.GetFullPath(root);
        if(extractOnly&&Directory.Exists(root))throw new IOException("检查目录必须尚不存在，以免覆盖文件。");
        Directory.CreateDirectory(root);VersionedInstallation.RejectReparse(root);
        using(var resource=Assembly.GetExecutingAssembly().GetManifestResourceStream("payload.zip"))
        using(var archive=new ZipArchive(resource,ZipArchiveMode.Read))
        foreach(var entry in archive.Entries)
        {
            string destination=Path.GetFullPath(Path.Combine(root,entry.FullName));
            if(!destination.StartsWith(root.TrimEnd(Path.DirectorySeparatorChar)+Path.DirectorySeparatorChar,StringComparison.OrdinalIgnoreCase))throw new InvalidDataException("安装文件路径无效。");
            Directory.CreateDirectory(Path.GetDirectoryName(destination));VersionedInstallation.RejectReparse(Path.GetDirectoryName(destination));
            if(entry.Name.Length==0)continue;
            if(File.Exists(destination))VersionedInstallation.RejectReparse(destination);
            using(var input=entry.Open())using(var output=File.Create(destination))input.CopyTo(output);
        }
    }
    static int RunMsi(string command)
    {
        using(var process=Process.Start(new ProcessStartInfo(Path.Combine(Environment.SystemDirectory,"msiexec.exe"),command+" EXCELNAV_SETUP=1 /passive /norestart REBOOT=ReallySuppress MSIRESTARTMANAGERCONTROL=Disable /l*v \""+diagnosticFile+".msi.log\""){UseShellExecute=false,CreateNoWindow=true}))
        {process.WaitForExit();return process.ExitCode;}
    }
    static void RequireMsiSuccess(int code,string action)
    {if(code!=0&&code!=3010&&code!=1641)throw new InvalidOperationException(action+"未完成，Windows Installer 返回代码："+code);}
    static void PrepareUserRegistration(VersionedInstallation installation,bool apply)
    {
        foreach(var view in new[]{RegistryView.Registry32,RegistryView.Registry64})
        {
            if(view==RegistryView.Registry64&&!Environment.Is64BitOperatingSystem)continue;
            using(var user=RegistryKey.OpenBaseKey(RegistryHive.CurrentUser,view))
            using(var key=user.OpenSubKey(VersionedInstallation.OfficeKey))
            {
                string manifest=Convert.ToString(key==null?null:key.GetValue("Manifest"));
                if(manifest.Length==0)continue;
                if(installation.IsOwnedManifest(manifest))
                {
                    if(apply)BackupAndRemoveRegistration(user,VersionedInstallation.OfficeKey,manifest);
                }
                else EnsureNoLegacyRegistration(!apply);
            }
        }
    }
    static void CleanupMaintenance(string root)
    {
        foreach(string folder in new[]{Path.Combine(root,"Versions"),Path.Combine(root,"Installer"),Path.Combine(root,"Maintenance")})
        {
            if(!Directory.Exists(folder))continue;
            VersionedInstallation.RejectReparse(folder);
            foreach(string file in Directory.GetFiles(folder,"*",SearchOption.AllDirectories))
            {
                VersionedInstallation.RejectReparse(file);
                // Never schedule delayed deletion: reinstalling before reboot must remain safe.
                try{File.Delete(file);}catch(IOException){ }
            }
            foreach(string directory in Directory.GetDirectories(folder,"*",SearchOption.AllDirectories).OrderByDescending(p=>p.Length))
                if(Directory.GetFileSystemEntries(directory).Length==0)Directory.Delete(directory);
            if(Directory.GetFileSystemEntries(folder).Length==0)Directory.Delete(folder);
        }
    }

    internal static void EnsureOfficeClosed(string reason = null)
    {
        WaitForOfficeClosed(IsOfficeRunning, () => MessageBox.Show(
            (reason==null?"":reason+"\n\n")+"请先保存所有工作，再关闭全部 Excel 和 WPS 表格窗口。\n\n关闭后点击“重试”继续安装；点击“取消”退出安装。",
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
        if (exitCode == 0) return "Excel Navigator 安装成功。\n\n当前 Excel/WPS 窗口可以继续使用；新的 Excel/WPS 进程将加载新版。全部退出后重新打开即可生效。";
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
        BackupAndRemoveRegistration(root,path,expectedManifest);
    }

    internal static void BackupAndRemoveRegistration(RegistryKey root,string path,string expectedManifest)
    {
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
