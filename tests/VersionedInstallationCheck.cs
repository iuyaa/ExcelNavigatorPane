// Controlled registry roots and temporary signed fixtures; never touches Office registration.
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Diagnostics;
using System.Xml;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using Microsoft.Win32;
using Microsoft.CSharp;
using System.CodeDom.Compiler;

class VersionedInstallationCheck
{
    static void Need(bool value,string reason) { if(!value)throw new Exception(reason); }
    static void Reject(Action action,string reason) { bool rejected=false;try{action();}catch{rejected=true;}Need(rejected,reason); }
    static void SignedManifest(string file,string name,string version,string dependency,RSA rsa)
    {
        string digest;
        using(var sha=SHA256.Create())digest=Convert.ToBase64String(sha.ComputeHash(File.ReadAllBytes(Path.Combine(Path.GetDirectoryName(file),dependency))));
        var xml=new XmlDocument { PreserveWhitespace=true };
        xml.LoadXml("<assembly><assemblyIdentity name='"+name+"' version='"+version+"'/><dependency><dependentAssembly codebase='"+dependency+"'><hash><DigestMethod Algorithm='http://www.w3.org/2001/04/xmlenc#sha256'/><DigestValue>"+digest+"</DigestValue></hash></dependentAssembly></dependency></assembly>");
        var signature=new SignedXml(xml) { SigningKey=rsa };
        signature.SignedInfo.SignatureMethod="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
        var reference=new Reference("") { DigestMethod="http://www.w3.org/2001/04/xmlenc#sha256" };
        reference.AddTransform(new XmlDsigEnvelopedSignatureTransform());signature.AddReference(reference);signature.ComputeSignature();
        xml.DocumentElement.AppendChild(xml.ImportNode(signature.GetXml(),true));xml.Save(file);
    }
    static void Fixture(string root,string version,RSA rsa)
    {
        string dir=Path.Combine(root,"Versions",version);Directory.CreateDirectory(dir);
        using(var provider=new CSharpCodeProvider())
        {
            var result=provider.CompileAssemblyFromSource(new CompilerParameters { GenerateExecutable=false,OutputAssembly=Path.Combine(dir,"ExcelNavigatorPane.dll") },
                "[assembly:System.Reflection.AssemblyVersion(\""+version+"\")][assembly:System.Reflection.AssemblyFileVersion(\""+version+"\")]public static class Payload{public static string Ping(){return \""+version+"\";}}");
            Need(!result.Errors.HasErrors,"Fixture compile failed");
        }
        SignedManifest(Path.Combine(dir,"ExcelNavigatorPane.dll.manifest"),"ExcelNavigatorPane.dll",version,"ExcelNavigatorPane.dll",rsa);
        SignedManifest(Path.Combine(dir,"ExcelNavigatorPane.vsto"),"ExcelNavigatorPane.vsto",version,"ExcelNavigatorPane.dll.manifest",rsa);
        Directory.CreateDirectory(Path.Combine(root,"Maintenance",version));
        File.Copy(Assembly.GetExecutingAssembly().Location,Path.Combine(root,"Maintenance",version,"ExcelNavigator-Setup.exe"));
    }
    static string Current(RegistryKey scope) { using(var key=scope.OpenSubKey(VersionedInstallation.StateKey))return Convert.ToString(key==null?null:key.GetValue("CurrentVersion")); }
    static string Manifest(RegistryKey scope) { using(var key=scope.OpenSubKey(VersionedInstallation.OfficeKey))return Convert.ToString(key==null?null:key.GetValue("Manifest")); }
    static void Crash(string root,string testKey,string keyFile,string version)
    {
        using(var a=Registry.CurrentUser.OpenSubKey(testKey+"\\a",true))
        using(var b=Registry.CurrentUser.OpenSubKey(testKey+"\\b",true))
        {
            var manager=new VersionedInstallation(root,File.ReadAllText(keyFile),a,b);
            manager.AfterWrite=i=>{if(i==4)Environment.Exit(71);};manager.Activate(version,false);
        }
    }
    static int Main(string[] args)
    {
        if(args.Length==5&&args[0]=="--crash") { Crash(args[1],args[2],args[3],args[4]);return 2; }
        string root=Path.Combine(Path.GetTempPath(),"ExcelNavigator-VersionedCheck-"+Guid.NewGuid().ToString("N"));
        string testKey=@"Software\ExcelNavigatorPane\Tests\Versioned\"+Guid.NewGuid().ToString("N");
        try
        {
            Directory.CreateDirectory(root);
            using(var test=Registry.CurrentUser.CreateSubKey(testKey))
            using(var a=test.CreateSubKey("a"))
            using(var b=test.CreateSubKey("b"))
            {
                string realVersion=args[2],actual=Path.Combine(root,"real"),realFolder=Path.Combine(actual,"Versions",realVersion);
                Directory.CreateDirectory(realFolder);
                foreach(var file in Directory.GetFiles(args[0]).Where(f=>!f.EndsWith(".wxs")))File.Copy(file,Path.Combine(realFolder,Path.GetFileName(file)),true);
                new VersionedInstallation(actual,File.ReadAllText(args[1]),a,b).ValidateVersion(realVersion);
                using(var wrong=RSA.Create())Reject(()=>new VersionedInstallation(actual,wrong.ToXmlString(false),a,b).ValidateVersion(realVersion),"Wrong signing key accepted");
                Console.WriteLine("PASS: packaged VSTO RSA signatures, identities, versions and dependency hashes verified; wrong key rejected");
                using(var rsa=RSA.Create())
                {
                    string old="1.0.26.0",next="1.0.27.0",key=rsa.ToXmlString(false),keyFile=Path.Combine(root,"public-key.xml");File.WriteAllText(keyFile,key);
                    Fixture(root,old,rsa);Fixture(root,next,rsa);Fixture(root,"1.0.28.0",rsa);
                    var manager=new VersionedInstallation(root,key,a,b);
                    foreach(string file in Directory.GetFiles(manager.VersionFolder(old)))File.Copy(file,Path.Combine(root,Path.GetFileName(file)));
                    using(var office=a.CreateSubKey(VersionedInstallation.OfficeKey))office.SetValue("Manifest",new Uri(Path.Combine(root,"ExcelNavigatorPane.vsto")).AbsoluteUri+"|vstolocal");
                    Need(manager.NeedsMigration(),"Legacy layout was not detected");
                    Need(manager.BackupLegacy(Assembly.GetExecutingAssembly().Location)==old,"Signed migration backup failed");
                    manager.StageMigration(old);
                    a.DeleteSubKeyTree(VersionedInstallation.OfficeKey,false);
                    manager.RecoverMigration(false);
                    Need(Current(a)==old&&manager.HasActiveRegistration(old)&&manager.ReadState("MigrationFallback")=="","Interrupted legacy removal lost signed fallback");
                    manager.Activate(old,false);Need(!manager.NeedsMigration()&&Current(a)==old,"First activation failed");
                    Need(manager.IsOwnedManifest(manager.Manifest(old))&&!manager.IsOwnedManifest(manager.Manifest(old).Replace("Versions","Other")),"Manifest ownership boundary failed");
                    string path=Path.Combine(manager.VersionFolder(old),"ExcelNavigatorPane.dll");
                    using(var held=new FileStream(path,FileMode.Open,FileAccess.Read,FileShare.Read))
                    {
                        var payload=Assembly.LoadFile(path).GetType("Payload").GetMethod("Ping");
                        Need((string)payload.Invoke(null,null)==old,"Old payload failed");
                        manager.Activate(next,false);
                        Need((string)payload.Invoke(null,null)==old&&Current(a)==next&&Manifest(a)==manager.Manifest(next)&&Manifest(b)==manager.Manifest(next),"Live activation changed old execution or missed a registry view");
                        Need((string)Assembly.LoadFile(Path.Combine(manager.VersionFolder(next),"ExcelNavigatorPane.dll")).GetType("Payload").GetMethod("Ping").Invoke(null,null)==next,"New load failed");
                    }
                    manager.Activate(old,true);Need(Current(a)==old&&manager.ReadState("PreviousVersion")==next,"Restore must retain both versions");
                    manager.AfterWrite=i=>{if(i==5)throw new IOException("Injected second-view failure");};
                    Reject(()=>manager.Activate(next,false),"Injected activation failure accepted");manager.AfterWrite=null;
                    Need(Current(a)==old&&Manifest(a)==manager.Manifest(old)&&Manifest(b)==manager.Manifest(old)&&!File.Exists(Path.Combine(root,"activation-journal.xml")),"Failed activation did not restore both views");
                    Console.WriteLine("PASS: legacy detection, old DLL still executes, new load, two-view promotion, explicit restore and partial-write rollback");
                    string command="--crash \""+root+"\" \""+testKey+"\" \""+keyFile+"\" "+next;
                    using(var child=Process.Start(new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,command) { UseShellExecute=false,CreateNoWindow=true }))
                        Need(child.WaitForExit(20000)&&child.ExitCode==71,"Crash injection failed");
                    Need(File.Exists(Path.Combine(root,"activation-journal.xml")),"Crash lost recovery journal");manager.RecoverActivation();
                    Need(Current(a)==old&&Manifest(a)==manager.Manifest(old)&&Manifest(b)==manager.Manifest(old),"Restart recovery failed");
                    using(var child=Process.Start(new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,command) { UseShellExecute=false,CreateNoWindow=true }))
                        Need(child.WaitForExit(20000)&&child.ExitCode==71,"Second crash injection failed");
                    using(var office=a.CreateSubKey(VersionedInstallation.OfficeKey))office.SetValue("Manifest","externally-changed");
                    Reject(()=>manager.RecoverActivation(),"Recovery overwrote a concurrent registry change");
                    Need(Manifest(a)=="externally-changed"&&File.Exists(Path.Combine(root,"activation-journal.xml")),"Conflict must retain journal and foreign value");
                    using(var office=a.CreateSubKey(VersionedInstallation.OfficeKey))office.SetValue("Manifest",manager.Manifest(next));
                    manager.RecoverActivation();
                    manager.Activate(next,false);Reject(()=>manager.Activate(old,false),"Unrequested downgrade accepted");
                    using(var office=b.CreateSubKey(VersionedInstallation.OfficeKey))office.SetValue("Manifest","file:///C:/foreign/addin.vsto|vstolocal");
                    Reject(()=>manager.Activate(next,false),"Activation overwrote foreign registration");
                    Reject(()=>manager.RemoveRegistration(),"Foreign registration removed");Need(Current(a)==next,"Foreign uninstall changed state");
                    using(var office=b.CreateSubKey(VersionedInstallation.OfficeKey))office.SetValue("Manifest",manager.Manifest(next));
                    manager.RemoveRegistration();Need(Current(a)==""&&Manifest(a)==""&&Manifest(b)=="","Unified registration cleanup failed");
                    string badVersion="1.0.29.0";Fixture(root,badVersion,rsa);
                    File.AppendAllText(Path.Combine(manager.VersionFolder(badVersion),"ExcelNavigatorPane.dll"),"tampered");
                    Reject(()=>manager.ValidateVersion(badVersion),"Tampered payload accepted");
                    foreach(string invalid in new[]{"../1.0.1.0","1.0.1","1.0.1.1","256.0.1.0","01.0.1.0"})Reject(()=>manager.VersionFolder(invalid),"Bad version path accepted");
                    Console.WriteLine("PASS: process-crash recovery, downgrade/tamper/path guards and foreign-registration-safe uninstall");
                }
            }
            return 0;
        }
        catch(Exception e) { Console.WriteLine(e);return 1; }
        finally
        {
            Registry.CurrentUser.DeleteSubKeyTree(testKey,false);
            Console.WriteLine("Evidence (temporary signed fixtures): "+root);
        }
    }
}
