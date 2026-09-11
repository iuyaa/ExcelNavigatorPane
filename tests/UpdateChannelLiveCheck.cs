// Compile with UpdateChecker.cs and a copy of packaged settings using an older probe version.
// Arguments: expected remote version, output EXE path, matching locally built EXE path. Never installs.
using System;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using ExcelNavigatorPane;

class UpdateChannelLiveCheck
{
    static int Main(string[] args)
    {
        try
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
            var release = UpdateChecker.CheckAsync().GetAwaiter().GetResult();
            if (release.Version == null || release.Version.ToString() != args[0]) throw new Exception(release.Message);
            UpdateChecker.DownloadAsync(release, args[1]).GetAwaiter().GetResult();
            using (var sha = SHA256.Create())
            {
                string expected = BitConverter.ToString(sha.ComputeHash(File.ReadAllBytes(args[2])));
                string actual = BitConverter.ToString(sha.ComputeHash(File.ReadAllBytes(args[1])));
                if (actual != expected) throw new Exception("Downloaded artifact differs from channel build");
                Console.WriteLine("SHA256: " + actual.Replace("-", "").ToLowerInvariant());
            }
            if (ServicePointManager.SecurityProtocol != SecurityProtocolType.Tls) throw new Exception("Global TLS changed");
            Console.WriteLine("PASS: " + UpdateChecker.ReadSettings().Attribute("channel") + "; " + release.DownloadUrl +
                "; live discovery, signature, download/hash, legacy TLS. Actual install NOT_RUN.");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex); return 1; }
    }
}
