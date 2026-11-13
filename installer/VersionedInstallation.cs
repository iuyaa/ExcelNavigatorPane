using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Win32;

// Only installation/maintenance executables use this class; Office never writes machine state.
internal sealed class VersionedInstallation
{
    internal const string PayloadFamily = "{4BCB8A50-9302-4E53-8829-F103215C070E}";
    internal const string OfficeKey = @"Software\Microsoft\Office\Excel\Addins\ExcelNavigatorPane";
    internal const string StateKey = @"Software\ExcelNavigatorPane\Installation";
    internal const string UninstallKey = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\ExcelNavigatorPane.Versioned";
    readonly string root, publicKey;
    readonly RegistryKey[] scopes;
    internal Action<int> AfterWrite = null;
    string Journal { get { return Path.Combine(root, "activation-journal.xml"); } }
    internal string Root { get { return root; } }

    internal VersionedInstallation(string root, string publicKey, params RegistryKey[] scopes)
    {
        this.root = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar);
        this.publicKey = publicKey;
        this.scopes = scopes;
    }
    internal static Version ParseVersion(string text)
    {
        Version v;
        if (!Version.TryParse(text, out v) || v.ToString() != text || v.Revision != 0 || v.Major > 255 || v.Minor > 255 || v.Build > 65535 || v < new Version(1,0,1,0))
            throw new InvalidDataException("安装版本号无效。");
        return v;
    }
    internal string VersionFolder(string version) { ParseVersion(version); return Path.Combine(root, "Versions", version); }
    internal string ManagerPath(string version) { ParseVersion(version); return Path.Combine(root, "Maintenance", version, "ExcelNavigator-Setup.exe"); }
    internal string Manifest(string version) { return new Uri(Path.Combine(VersionFolder(version), "ExcelNavigatorPane.vsto")).AbsoluteUri + "|vstolocal"; }
    internal string ReadState(string name)
    {
        using (var key = scopes[0].OpenSubKey(StateKey)) return Convert.ToString(key == null ? null : key.GetValue(name));
    }
    internal bool NeedsMigration()
    {
        foreach (var scope in scopes)
        using (var key = scope.OpenSubKey(OfficeKey))
        {
            string manifest = Convert.ToString(key == null ? null : key.GetValue("Manifest"));
            if (manifest.Length == 0) continue;
            if (!IsOwnedManifest(manifest)) throw new InvalidOperationException("加载项注册指向另一份安装，请先处理注册冲突。");
            if (!manifest.Contains("/Versions/")) return true;
        }
        return false;
    }
    internal bool IsOwnedManifest(string manifest)
    {
        const string suffix = "|vstolocal";
        Uri uri;
        if (manifest == null || !manifest.EndsWith(suffix, StringComparison.OrdinalIgnoreCase) ||
            !Uri.TryCreate(manifest.Substring(0, manifest.Length-suffix.Length), UriKind.Absolute, out uri) ||
            !uri.IsFile || uri.IsUnc || uri.Query.Length != 0 || uri.Fragment.Length != 0) return false;
        string path = Path.GetFullPath(uri.LocalPath);
        if (string.Equals(path, Path.Combine(root, "ExcelNavigatorPane.vsto"), StringComparison.OrdinalIgnoreCase)) return true;
        string folder = Path.GetDirectoryName(path);
        Version version;
        return Path.GetFileName(path).Equals("ExcelNavigatorPane.vsto", StringComparison.OrdinalIgnoreCase) &&
            Version.TryParse(Path.GetFileName(folder),out version) && version.Revision==0 && version.ToString()==Path.GetFileName(folder) &&
            string.Equals(folder, VersionFolder(version.ToString()), StringComparison.OrdinalIgnoreCase);
    }
    internal void ValidateVersion(string version)
    {
        string folder = VersionFolder(version);
        RejectReparse(folder);
        var deployment = VerifyManifest(Path.Combine(folder,"ExcelNavigatorPane.vsto"),version);
        var application = VerifyManifest(Path.Combine(folder,"ExcelNavigatorPane.dll.manifest"),version);
        foreach (var document in new[] { deployment, application })
        foreach (XmlElement dependency in document.SelectNodes("//*[local-name()='dependentAssembly'][@codebase]"))
        {
            string name = dependency.GetAttribute("codebase");
            if (string.IsNullOrEmpty(name) || Path.GetFileName(name) != name || name.IndexOfAny(new[]{'/', '\\', ':'}) >= 0 || name == "." || name == "..")
                throw new InvalidDataException("清单包含非本地文件。");
            string file = Path.Combine(folder,name);
            RejectReparse(file);
            var digest = dependency.SelectSingleNode(".//*[local-name()='DigestValue']");
            var method = (XmlElement)dependency.SelectSingleNode(".//*[local-name()='DigestMethod']");
            if (digest == null || method == null || !method.GetAttribute("Algorithm").EndsWith("#sha256",StringComparison.Ordinal))
                throw new InvalidDataException("清单缺少 SHA256 文件校验。");
            using (var sha=SHA256.Create())
            using (var stream=File.OpenRead(file))
                if (Convert.ToBase64String(sha.ComputeHash(stream)) != digest.InnerText) throw new InvalidDataException("已安装文件校验失败："+name);
        }
        if (FileVersionInfo.GetVersionInfo(Path.Combine(folder,"ExcelNavigatorPane.dll")).FileVersion != version)
            throw new InvalidDataException("插件文件版本不一致。");
    }
    XmlDocument VerifyManifest(string file,string version)
    {
        RejectReparse(file);
        var doc=new XmlDocument { PreserveWhitespace=true,XmlResolver=null };
        using(var reader=XmlReader.Create(file,new XmlReaderSettings { DtdProcessing=DtdProcessing.Prohibit, XmlResolver=null,MaxCharactersInDocument=4*1024*1024 }))doc.Load(reader);
        var identity=(XmlElement)doc.DocumentElement.SelectSingleNode("*[local-name()='assemblyIdentity']");
        string expectedName=Path.GetFileName(file)=="ExcelNavigatorPane.vsto"?"ExcelNavigatorPane.vsto":"ExcelNavigatorPane.dll";
        if(identity==null || identity.GetAttribute("version")!=version || identity.GetAttribute("name")!=expectedName)throw new InvalidDataException("清单身份或版本不一致。");
        var signatures=doc.DocumentElement.SelectNodes("*[local-name()='Signature' and namespace-uri()='http://www.w3.org/2000/09/xmldsig#']");
        if(signatures.Count!=1)throw new InvalidDataException("清单签名不唯一。");
        CryptoConfig.AddAlgorithm(typeof(VstoSha256Signature),"http://www.w3.org/2000/09/xmldsig#rsa-sha256");
        CryptoConfig.AddAlgorithm(typeof(SHA256Managed),"http://www.w3.org/2000/09/xmldsig#sha256");
        var signed=new SignedXml(doc);signed.LoadXml((XmlElement)signatures[0]);
        if(signed.SignedInfo.SignatureMethod!="http://www.w3.org/2000/09/xmldsig#rsa-sha256" &&
            signed.SignedInfo.SignatureMethod!="http://www.w3.org/2001/04/xmldsig-more#rsa-sha256")throw new InvalidDataException("清单必须使用 RSA-SHA256 签名。");
        bool wholeDocument=false;
        foreach(Reference reference in signed.SignedInfo.References)
        {
            if(reference.Uri!="" && reference.Uri!="#StrongNameKeyInfo")throw new InvalidDataException("清单签名引用无效。");
            if(reference.Uri=="")wholeDocument=true;
        }
        if(!wholeDocument)throw new InvalidDataException("清单签名未覆盖完整文档。");
        using(var rsa=RSA.Create()) { rsa.FromXmlString(publicKey);if(!signed.CheckSignature(rsa))throw new InvalidDataException("清单签名与发布密钥不匹配。"); }
        return doc;
    }
    internal static void RejectReparse(string path)
    {
        for(string current=Path.GetFullPath(path);!string.IsNullOrEmpty(current);current=Path.GetDirectoryName(current))
            if((File.GetAttributes(current)&FileAttributes.ReparsePoint)!=0)throw new InvalidDataException("安装路径不能包含重解析点。");
    }
    internal void StoreManager(string version,string source)
    {
        string target=ManagerPath(version);
        Directory.CreateDirectory(Path.GetDirectoryName(target));RejectReparse(Path.GetDirectoryName(target));
        if(File.Exists(target))
        {
            if(Hash(target)!=Hash(source))throw new InvalidOperationException("同一版本的安装器内容不同，请使用递增版本。");
        }
        else File.Copy(source,target,false);
    }
    static string Hash(string path) { using(var sha=SHA256.Create())using(var file=File.OpenRead(path))return Convert.ToBase64String(sha.ComputeHash(file)); }

    internal string BackupLegacy(string managerSource)
    {
        if(!NeedsMigration())return null;
        string version=FileVersionInfo.GetVersionInfo(Path.Combine(root,"ExcelNavigatorPane.dll")).FileVersion;
        string destination=VersionFolder(version);Directory.CreateDirectory(destination);RejectReparse(destination);
        foreach(string file in Directory.GetFiles(root).Where(f=>new[]{".dll",".manifest",".vsto",".config"}.Contains(Path.GetExtension(f).ToLowerInvariant())))
        {
            RejectReparse(file);string target=Path.Combine(destination,Path.GetFileName(file));
            if(File.Exists(target)){RejectReparse(target);if(Hash(target)!=Hash(file))throw new InvalidDataException("迁移备份与已保留版本内容不同。");}
            else File.Copy(file,target,false);
        }
        ValidateVersion(version);StoreManager(version,managerSource);return version;
    }
    internal void StageMigration(string version)
    {
        ValidateVersion(version);
        using(var key=scopes[0].CreateSubKey(StateKey)){key.SetValue("MigrationFallback",version,RegistryValueKind.String);key.Flush();}
    }
    internal void ClearMigration()
    {
        using(var key=scopes[0].OpenSubKey(StateKey,true))if(key!=null){key.DeleteValue("MigrationFallback",false);key.Flush();}
    }
    internal bool HasActiveRegistration(string version)
    {
        foreach(var scope in scopes)
        using(var key=scope.OpenSubKey(OfficeKey))
            if(key==null||!string.Equals(Convert.ToString(key.GetValue("Manifest")),Manifest(version),StringComparison.OrdinalIgnoreCase))return false;
        return true;
    }
    internal void RecoverMigration(bool legacyPresent)
    {
        string fallback=ReadState("MigrationFallback");if(fallback.Length==0)return;
        ValidateVersion(fallback);
        string current=ReadState("CurrentVersion");
        if(!legacyPresent)
        {
            if(current.Length==0 || current==fallback)Activate(fallback,true);
            else { ValidateVersion(current);if(!HasActiveRegistration(current))throw new InvalidOperationException("迁移恢复时加载入口不一致，保留恢复记录。"); }
        }
        ClearMigration();
    }

    internal void Activate(string version,bool restore,string migratedPrevious = null)
    {
        EnsureOwnedRegistration();ValidateVersion(version);
        if(!File.Exists(ManagerPath(version)))throw new FileNotFoundException("缺少维护程序，不能切换版本。");
        string previous=ReadState("CurrentVersion");
        if(previous.Length==0 && migratedPrevious!=null){ValidateVersion(migratedPrevious);previous=migratedPrevious;}
        if(!restore && previous.Length!=0 && ParseVersion(previous)>ParseVersion(version))throw new InvalidOperationException("不能覆盖较新版本；如需恢复请使用恢复上一版。");
        var changes=new List<Change>();
        for(int i=0;i<scopes.Length;i++)
        {
            Add(changes,i,OfficeKey,"Manifest",Manifest(version));
            Add(changes,i,OfficeKey,"FriendlyName","Excel Navigator");
            Add(changes,i,OfficeKey,"Description","工作簿和工作表导航");
            Add(changes,i,OfficeKey,"LoadBehavior",3);
        }
        if(previous!=version && previous.Length!=0)Add(changes,0,StateKey,"PreviousVersion",previous);
        Add(changes,0,StateKey,"CurrentVersion",version);
        Add(changes,0,UninstallKey,"DisplayName","Excel Navigator");
        Add(changes,0,UninstallKey,"DisplayVersion",version);
        Add(changes,0,UninstallKey,"Publisher","Excel Navigator");
        Add(changes,0,UninstallKey,"InstallLocation",root);
        Add(changes,0,UninstallKey,"UninstallString","\""+ManagerPath(version)+"\" --uninstall");
        Add(changes,0,UninstallKey,"ModifyPath","\""+ManagerPath(version)+"\" --restore");
        Add(changes,0,UninstallKey,"NoRepair",1);
        var xml=new XElement("activation",changes.Select(c=>c.ToXml()));
        Directory.CreateDirectory(root);RejectReparse(root);
        using(var stream=new FileStream(Journal,FileMode.CreateNew,FileAccess.Write,FileShare.None)) { xml.Save(stream);stream.Flush(true); }
        try
        {
            for(int i=0;i<changes.Count;i++){Write(changes[i],false);if(AfterWrite!=null)AfterWrite(i);}
            File.Delete(Journal);
        }
        catch { RecoverActivation();throw; }
    }
    sealed class Change
    {
        internal int Scope;internal string Key,Name;internal object Before,After;
        internal XElement ToXml(){return new XElement("value",new XAttribute("scope",Scope),new XAttribute("key",Key),new XAttribute("name",Name),new XAttribute("integer",After is int),new XAttribute("missing",Before==null),new XAttribute("before",Before??""),new XAttribute("after",After));}
    }
    void Add(List<Change> list,int scope,string key,string name,object value)
    {
        using(var old=scopes[scope].OpenSubKey(key))
        {
            object before=old==null?null:old.GetValue(name,null,RegistryValueOptions.DoNotExpandEnvironmentNames);
            var expected=value is int?RegistryValueKind.DWord:RegistryValueKind.String;
            if(before!=null && old.GetValueKind(name)!=expected)throw new InvalidDataException("注册值类型异常："+name);
            list.Add(new Change{Scope=scope,Key=key,Name=name,Before=before,After=value});
        }
    }
    void Write(Change change,bool rollback)
    {
        object value=rollback?change.Before:change.After;
        using(var key=scopes[change.Scope].CreateSubKey(change.Key))
        { if(value==null)key.DeleteValue(change.Name,false);else key.SetValue(change.Name,value,value is int?RegistryValueKind.DWord:RegistryValueKind.String);key.Flush(); }
    }
    internal void RecoverActivation()
    {
        if(!File.Exists(Journal))return;
        XElement xml;
        using(var reader=XmlReader.Create(Journal,new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=128*1024}))xml=XElement.Load(reader);
        var changes=new List<Change>();var unique=new HashSet<string>();
        foreach(var e in xml.Elements("value"))
        {
            int scope=(int)e.Attribute("scope");string key=(string)e.Attribute("key"),name=(string)e.Attribute("name");
            string names=key==OfficeKey?"Manifest|FriendlyName|Description|LoadBehavior":key==StateKey?"CurrentVersion|PreviousVersion":key==UninstallKey?"DisplayName|DisplayVersion|Publisher|InstallLocation|UninstallString|ModifyPath|NoRepair":"";
            if(scope<0||scope>=scopes.Length||!names.Split('|').Contains(name)||!unique.Add(scope+"|"+key+"|"+name))throw new InvalidDataException("版本切换恢复记录无效。");
            bool integer=(bool)e.Attribute("integer"),missing=(bool)e.Attribute("missing");
            object before=missing?null:integer?(object)int.Parse((string)e.Attribute("before")):(string)e.Attribute("before");
            object after=integer?(object)int.Parse((string)e.Attribute("after")):(string)e.Attribute("after");
            using(var current=scopes[scope].OpenSubKey(key))
            {
                object value=current==null?null:current.GetValue(name);
                if(!object.Equals(value,before)&&!object.Equals(value,after))throw new InvalidOperationException("版本切换后注册被其他程序更改，保留恢复记录，请联系维护人员。");
            }
            changes.Add(new Change{Scope=scope,Key=key,Name=name,Before=before,After=after});
        }
        if(changes.Count==0)throw new InvalidDataException("版本切换恢复记录为空。");
        for(int i=changes.Count-1;i>=0;i--)Write(changes[i],true);
        File.Delete(Journal);
    }
    internal void EnsureOwnedRegistration()
    {
        RecoverActivation();
        foreach(var scope in scopes)
        using(var key=scope.OpenSubKey(OfficeKey))
        {
            if(key!=null&&!IsOwnedManifest(Convert.ToString(key.GetValue("Manifest"))))throw new InvalidOperationException("注册已被其他安装更改，停止卸载。");
        }
    }
    internal void RemoveRegistration()
    {
        EnsureOwnedRegistration();
        foreach(var scope in scopes)scope.DeleteSubKeyTree(OfficeKey,false);
        scopes[0].DeleteSubKeyTree(StateKey,false);scopes[0].DeleteSubKeyTree(UninstallKey,false);
    }
    internal static List<string> Products(string family)
    {
        var list=new List<string>();
        for(uint i=0;;i++){var code=new StringBuilder(39);uint result=MsiEnumRelatedProducts(family,0,i,code);if(result==259)break;if(result!=0)throw new System.ComponentModel.Win32Exception((int)result);list.Add(code.ToString());}
        return list;
    }
    internal static string ProductVersion(string product)
    {
        uint size=128;var text=new StringBuilder((int)size);uint result=MsiGetProductInfo(product,"VersionString",text,ref size);
        if(result!=0)throw new System.ComponentModel.Win32Exception((int)result);return text.ToString();
    }
    [DllImport("msi.dll",CharSet=CharSet.Unicode)]static extern uint MsiEnumRelatedProducts(string family,uint reserved,uint index,StringBuilder product);
    [DllImport("msi.dll",CharSet=CharSet.Unicode)]static extern uint MsiGetProductInfo(string product,string property,StringBuilder value,ref uint size);
}

public sealed class VstoSha256Signature : SignatureDescription
{
    public VstoSha256Signature(){KeyAlgorithm=typeof(RSA).AssemblyQualifiedName;DigestAlgorithm=typeof(SHA256Managed).AssemblyQualifiedName;}
    public override AsymmetricSignatureFormatter CreateFormatter(AsymmetricAlgorithm key){var f=new RSAPKCS1SignatureFormatter(key);f.SetHashAlgorithm("SHA256");return f;}
    public override AsymmetricSignatureDeformatter CreateDeformatter(AsymmetricAlgorithm key){var f=new RSAPKCS1SignatureDeformatter(key);f.SetHashAlgorithm("SHA256");return f;}
}
