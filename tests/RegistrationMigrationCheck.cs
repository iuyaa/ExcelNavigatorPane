// Compile with installer/SetupLauncher.cs, /main:RegistrationMigrationCheck.
// Uses a unique HKCU test subtree, never the Office registration.
using System;
using Microsoft.Win32;

class RegistrationMigrationCheck
{
    static void Require(bool value) { if (!value) throw new Exception("Registration migration check failed"); }
    static int Main()
    {
        string testPath = @"Software\ExcelNavigatorPane\Tests\" + Guid.NewGuid().ToString("N");
        const string manifest = "file:///M:/project/bin/Debug/ExcelNavigatorPane.vsto|vstolocal";
        try
        {
            Require(SetupLauncher.IsDevelopmentRegistration(manifest));
            Require(SetupLauncher.IsDevelopmentRegistration(manifest.Replace("Debug", "Release")));
            foreach (string other in new[] { "https://host/bin/Debug/ExcelNavigatorPane.vsto|vstolocal", manifest.Replace("ExcelNavigatorPane.vsto", "Other.vsto"), manifest.Replace("bin/Debug", "Program Files/ExcelNavigatorPane"), manifest.Replace("file:///M:/", "file://server/share/"), "bad" })
                Require(!SetupLauncher.IsDevelopmentRegistration(other));
            using (var root = Registry.CurrentUser.CreateSubKey(testPath))
            {
                using (var key = root.CreateSubKey("Registration"))
                {
                    key.SetValue("Manifest", manifest);
                    key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                    key.SetValue("Expandable", "%TEMP%", RegistryValueKind.ExpandString);
                    key.SetValue("Binary", new byte[] { 1, 2 }, RegistryValueKind.Binary);
                    using (var child = key.CreateSubKey("Child")) child.SetValue("Text", "preserved");
                }
                bool rejected = false;
                try { SetupLauncher.BackupAndRemoveDevelopmentRegistration(root, "Registration", manifest.Replace("project", "changed")); }
                catch (InvalidOperationException) { rejected = true; }
                using (var retained = root.OpenSubKey("Registration")) Require(rejected && retained != null);
                SetupLauncher.BackupAndRemoveDevelopmentRegistration(root, "Registration", manifest);
                using (var removed = root.OpenSubKey("Registration")) Require(removed == null);
                using (var backups = root.OpenSubKey(@"Software\ExcelNavigatorPane\RegistrationBackups"))
                {
                    Require(backups.SubKeyCount == 1);
                    using (var backup = backups.OpenSubKey(backups.GetSubKeyNames()[0]))
                    {
                        Require((string)backup.GetValue("Manifest") == manifest);
                        Require(backup.GetValueKind("LoadBehavior") == RegistryValueKind.DWord && (int)backup.GetValue("LoadBehavior") == 3);
                        Require(backup.GetValueKind("Expandable") == RegistryValueKind.ExpandString && (string)backup.GetValue("Expandable", null, RegistryValueOptions.DoNotExpandEnvironmentNames) == "%TEMP%");
                        Require(((byte[])backup.GetValue("Binary"))[1] == 2);
                        using (var child = backup.OpenSubKey("Child")) Require((string)child.GetValue("Text") == "preserved");
                    }
                }
            }
            Console.WriteLine("PASS: development path identification, changed registration rejection, typed recursive backup and scoped removal (isolated test key only)");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex); return 1; }
        finally { Registry.CurrentUser.DeleteSubKeyTree(testPath, false); }
    }
}
