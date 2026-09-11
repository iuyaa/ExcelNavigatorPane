// Compile with UpdateChecker.cs, installer/SetupLauncher.cs and development settings.
// UpdateCheck.exe <release directory> <packaged ExcelNavigatorPane.dll>
using System;
using System.IO;
using System.Reflection;
using System.Net;
using System.Net.Sockets;
using System.Net.Http;
using System.Threading;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using ExcelNavigatorPane;

class UpdateCheck
{
    class RedirectHandler : HttpMessageHandler
    {
        internal string Location;
        internal int Calls;
        internal bool Repeat;
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken token)
        {
            var response = new HttpResponseMessage(++Calls == 1 || Repeat ? HttpStatusCode.Redirect : HttpStatusCode.OK);
            if (response.StatusCode == HttpStatusCode.Redirect) response.Headers.Location = new Uri(Location);
            return Task.FromResult(response);
        }
    }

    static void CheckChannels()
    {
        var address = UpdateChecker.ParseAddress("https://github.com/example/releases-test/releases.atom", "github");
        string feed = "<feed xmlns='http://www.w3.org/2005/Atom'>" +
            "<entry><link rel='alternate' href='https://github.com/example/releases-test/releases/tag/v1.0.9.0'/></entry>" +
            "<entry><link rel='alternate' href='https://github.com/example/releases-test/releases/tag/v1.0.10.0'/></entry>" +
            "<entry><link rel='alternate' href='https://evil.invalid/releases/tag/v99.0.0.0'/></entry></feed>";
        Require(UpdateChecker.GitHubManifest(feed, address).AbsoluteUri ==
            "https://github.com/example/releases-test/releases/download/v1.0.10.0/latest.xml");
        Reject(() => UpdateChecker.GitHubManifest("<feed xmlns='http://www.w3.org/2005/Atom'/>", address));
        Reject(() => UpdateChecker.GitHubManifest("<!DOCTYPE feed [<!ENTITY x SYSTEM 'file:///secret'>]><feed>&x;</feed>", address));
        Reject(() => UpdateChecker.ParseAddress(address.AbsoluteUri));
        Reject(() => UpdateChecker.ParseAddress("https://evil.invalid/example/repo/releases.atom", "github"));
        Reject(() => UpdateChecker.ParseAddress("", "unknown"));
        foreach (var target in new[] { "https://evil.invalid/file", "http://github.com/file", "https://user@github.com/file" })
            Require(!UpdateChecker.IsGitHubDownloadAddress(new Uri(target)));
        foreach (bool github in new[] { false, true })
        {
            var handler = new RedirectHandler { Location = "https://release-assets.githubusercontent.com/asset?signature=test" };
            using (var client = new HttpClient(handler))
            {
                Action request = () => { using (var response = UpdateChecker.GetResponseAsync(client, address, github).GetAwaiter().GetResult()) Require(response.IsSuccessStatusCode); };
                if (github) { request(); Require(handler.Calls == 2); } else Reject(request);
            }
        }
        foreach (var target in new[] { "https://evil.invalid/file", "https://github.com/loop" })
            using (var client = new HttpClient(new RedirectHandler { Location = target, Repeat = true }))
                Reject(() => UpdateChecker.GetResponseAsync(client, address, true).GetAwaiter().GetResult());
    }
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
            int officeChecks = 0, prompts = 0;
            CheckChannels();
            SetupLauncher.WaitForOfficeClosed(() => ++officeChecks <= 2,
                () => { prompts++; return System.Windows.Forms.DialogResult.Retry; });
            Require(officeChecks == 3 && prompts == 2);
            SetupLauncher.WaitForOfficeClosed(() => false,
                () => { throw new Exception("Closed Office must not prompt"); });
            bool cancelled = false;
            try { SetupLauncher.WaitForOfficeClosed(() => true, () => System.Windows.Forms.DialogResult.Cancel); }
            catch (OperationCanceledException) { cancelled = true; }
            Require(cancelled);
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
                string channel = args.Length > 2 ? args[2] : "oneview";
                Require(((string)settings.Attribute("channel") ?? "oneview") == channel);
                Require(UpdateChecker.ParseAddress((string)settings.Attribute("manifestUrl"), channel) != null);
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
            Console.WriteLine("PASS: channel/feed/redirect boundaries, Office wait/retry/cancel, signed release, pinned package key/version, newer version, tamper/wrong key/DTD/URL/hash rejection, disabled updates, Excel PE architecture, download/replace/cleanup (loopback only)");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex); return 1; }
    }
}
