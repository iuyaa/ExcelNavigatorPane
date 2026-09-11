// Compile with UpdateChecker.cs, installer/SetupLauncher.cs and development settings.
// UpdateCheck.exe <release directory> <packaged ExcelNavigatorPane.dll>
using System;
using System.IO;
using System.Reflection;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using ExcelNavigatorPane;

class UpdateCheck
{
    static void Require(bool value) { if (!value) throw new Exception("Check failed"); }
    static void Reject(Action action)
    {
        try { action(); }
        catch (InvalidDataException) { return; }
        catch (System.Xml.XmlException) { return; }
        throw new Exception("Invalid input accepted");
    }
    static void CheckDownload(string destination, bool valid)
    {
        byte[] body = Encoding.ASCII.GetBytes("download fixture");
        var listener = new TcpListener(IPAddress.Loopback, 0);
        listener.Start();
        var server = Task.Run(() =>
        {
            using (var client = listener.AcceptTcpClient())
            using (var stream = client.GetStream())
            using (var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, true))
            {
                string line;
                while (!string.IsNullOrEmpty(line = reader.ReadLine())) { }
                byte[] headers = Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\nContent-Length: " + body.Length + "\r\nConnection: close\r\n\r\n");
                stream.Write(headers,0,headers.Length); stream.Write(body,0,body.Length);
            }
        });
        try
        {
            string hash;
            using (var sha = SHA256.Create()) { hash = BitConverter.ToString(sha.ComputeHash(body)).Replace("-", "").ToLowerInvariant(); }
            // HTTP loopback is a test fixture only; production addresses come from the signed HTTPS parser.
            var release = new UpdateChecker.UpdateInfo { DownloadUrl = new Uri("http://127.0.0.1:" + ((IPEndPoint)listener.LocalEndpoint).Port + "/installer"), Sha256 = valid ? hash : new string('0',64) };
            File.WriteAllText(destination,"original file");
            Action download = () => UpdateChecker.DownloadAsync(release,destination).GetAwaiter().GetResult();
            if (valid) download(); else Reject(download);
            Require(File.ReadAllText(destination) == (valid ? "download fixture" : "original file"));
            Require(Directory.GetFiles(Path.GetDirectoryName(destination),Path.GetFileName(destination) + ".*.part").Length == 0);
            server.GetAwaiter().GetResult();
        }
        finally { listener.Stop(); }
    }
    static int Main(string[] args)
    {
        try
        {
            const string installedManifest = "file:///C:/Program Files/ExcelNavigatorPane/ExcelNavigatorPane.vsto|vstolocal";
            Require(!SetupLauncher.HasConflictingRegistration(null, installedManifest));
            Require(!SetupLauncher.HasConflictingRegistration("", installedManifest));
            Require(!SetupLauncher.HasConflictingRegistration(installedManifest, installedManifest));
            Require(!SetupLauncher.HasConflictingRegistration("file:///c:/program%20files/ExcelNavigatorPane/ExcelNavigatorPane.vsto|vstolocal", installedManifest));
            Require(!SetupLauncher.HasConflictingRegistration(@"file:///C:\Program Files\ExcelNavigatorPane\ExcelNavigatorPane.vsto|vstolocal", installedManifest));
            foreach (var conflict in new[] { "https://example.invalid/ExcelNavigatorPane.vsto", "file:///M:/project/bin/ExcelNavigatorPane.vsto|vstolocal", "not-a-manifest", "file:///C:/Program Files/ExcelNavigatorPane/ExcelNavigatorPane.vsto" })
                Require(SetupLauncher.HasConflictingRegistration(conflict, installedManifest));
            Require(SetupLauncher.HasConflictingRegistration(installedManifest, null));
            string manifest = File.ReadAllText(Path.Combine(args[0], "latest.xml"));
            string key = File.ReadAllText(Path.Combine(args[0], "update-public-key.xml")).Trim();
            var address = UpdateChecker.ParseAddress("https://updates.example.invalid/excel/latest.xml");
            Require(UpdateChecker.ParseAddress("") == null);
            foreach (var bad in new[] { "http://host/latest.xml", "https://user:secret@host/latest.xml", "https://host/",
                "https://host/latest.xml?token=secret", "file:///C:/latest.xml", "https://host/latest.xml#fragment" })
                Reject(() => UpdateChecker.ParseAddress(bad));
            var release = UpdateChecker.ReadRelease(manifest, address, key);
            Require(release.DownloadUrl.AbsoluteUri == "https://updates.example.invalid/excel/releases/" + release.Version + "/ExcelNavigator-Setup-" + release.Version + ".exe");
            UpdateChecker.VerifyInstaller(Path.Combine(args[0], Path.GetFileName(release.DownloadUrl.LocalPath)), release.Sha256);
            foreach (var edit in new[] { new[] { "product", "Other" }, new[] { "version", "1.0.1.1" },
                new[] { "version", "256.0.0.0" }, new[] { "file", "https://evil.invalid/installer.exe" },
                new[] { "file", "../installer.exe" }, new[] { "sha256", new string('0',64) }, new[] { "signature", "!" } })
            {
                var changed = XElement.Parse(manifest); changed.SetAttributeValue(edit[0], edit[1]);
                Reject(() => UpdateChecker.ReadRelease(changed.ToString(), address, key));
            }
            foreach (string bad in new[] { "<html />", "<!DOCTYPE release [<!ENTITY x SYSTEM 'file:///secret'>]><release>&x;</release>", new string('x',1048577) })
                Reject(() => UpdateChecker.ReadRelease(bad, address, key));
            using (var rsa = RSA.Create())
            {
                Reject(() => UpdateChecker.ReadRelease(manifest, address, rsa.ToXmlString(false)));
                var newer = XElement.Parse(manifest);
                string next = new Version(release.Version.Major, release.Version.Minor, release.Version.Build + 1, 0).ToString();
                newer.SetAttributeValue("version", next);
                newer.SetAttributeValue("file", "releases/" + next + "/ExcelNavigator-Setup-" + next + ".exe");
                var message = string.Join("\n", new[] { "ExcelNavigatorPane", next, (string)newer.Attribute("file"), release.Sha256 });
                newer.SetAttributeValue("signature", Convert.ToBase64String(rsa.SignData(Encoding.UTF8.GetBytes(message), HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1)));
                Require(UpdateChecker.ReadRelease(newer.ToString(), address, rsa.ToXmlString(false)).Version > release.Version);
            }
            Require(UpdateChecker.CheckAsync().GetAwaiter().GetResult().Message.Contains("尚未配置更新地址"));
            using (var stream = Assembly.LoadFrom(args[1]).GetManifestResourceStream("ExcelNavigatorPane.UpdateSettings.xml"))
            {
                var settings = XElement.Load(stream);
                Require((string)settings.Element("publicKey") == key);
                Require((string)settings.Attribute("version") == release.Version.ToString());
                Require(UpdateChecker.ParseAddress((string)settings.Attribute("manifestUrl")) != null);
            }
            string temp = Path.GetTempFileName();
            try
            {
                File.WriteAllText(temp,"tampered installer");
                Reject(() => UpdateChecker.VerifyInstaller(temp,release.Sha256));
                foreach (var machine in new ushort[] { 0x014c, 0x8664 })
                {
                    using (var writer = new BinaryWriter(File.Create(temp)))
                    {
                        writer.Write((ushort)0x5a4d); writer.BaseStream.Position = 0x3c; writer.Write(0x40);
                        writer.Write(0x00004550); writer.Write(machine);
                    }
                    Require(SetupLauncher.ReadPeArchitecture(temp) == (machine == 0x014c ? "x86" : "x64"));
                    Require(SetupLauncher.SelectInstallerArchitecture(true) == "x64");
                    Require(SetupLauncher.SelectInstallerArchitecture(false) == "x86");
                }
                File.WriteAllBytes(temp,new byte[128]); Reject(() => SetupLauncher.ReadPeArchitecture(temp));
                CheckDownload(temp,false); CheckDownload(temp,true);
            }
            finally { File.Delete(temp); }
            Console.WriteLine("PASS: signed release, pinned package key/version, newer version, tamper/wrong key/DTD/URL/hash rejection, disabled updates, Excel PE architecture, download/replace/cleanup (loopback only)");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex); return 1; }
    }
}
