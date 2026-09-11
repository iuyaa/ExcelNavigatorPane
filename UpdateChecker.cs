using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Authentication;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace ExcelNavigatorPane
{
    internal static class UpdateChecker
    {
        internal static HttpClientHandler CreateHttpHandler()
        {
            // VSTO hosts can retain legacy TLS defaults. Scope TLS to these requests, not Excel/WPS globally.
            return new HttpClientHandler { AllowAutoRedirect = false, SslProtocols = SslProtocols.Tls12 };
        }

        internal sealed class UpdateInfo
        {
            internal Version Version;
            internal Uri DownloadUrl;
            internal string Sha256;
            internal string Message;
            internal bool GitHub;
        }

        internal static XElement ReadSettings()
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelNavigatorPane.UpdateSettings.xml"))
                return XElement.Load(stream);
        }

        internal static Version CurrentVersion => Version.Parse((string)ReadSettings().Attribute("version"));

        internal static Uri ParseAddress(string address, string channel = "oneview")
        {
            if (channel != "oneview" && channel != "github") throw new InvalidDataException("未知的更新渠道。");
            if (string.IsNullOrWhiteSpace(address)) return null;
            Uri uri;
            if (!Uri.TryCreate(address, UriKind.Absolute, out uri) || uri.Scheme != Uri.UriSchemeHttps ||
                uri.UserInfo.Length != 0 || uri.Query.Length != 0 || uri.Fragment.Length != 0 ||
                (channel == "github" ? uri.Host != "github.com" ||
                    !Regex.IsMatch(uri.AbsolutePath, "\\A/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+/releases\\.atom\\z") :
                    !uri.AbsolutePath.EndsWith("/latest.xml", StringComparison.Ordinal)))
                throw new InvalidDataException("更新地址与渠道不匹配，必须使用固定的 HTTPS 地址。");
            return uri;
        }

        internal static Uri GitHubRepository(Uri feed)
        {
            return new Uri(feed, ".");
        }

        internal static Uri GitHubManifest(string xml, Uri address)
        {
            XElement feed = ReadXml(xml);
            XNamespace atom = "http://www.w3.org/2005/Atom";
            if (feed.Name != atom + "feed") throw new InvalidDataException("GitHub 发布列表格式不正确。");
            string prefix = new Uri(GitHubRepository(address), "releases/tag/v").AbsoluteUri;
            var versions = feed.Elements(atom + "entry").Elements(atom + "link")
                .Where(link => (string)link.Attribute("rel") == "alternate")
                .Select(link => (string)link.Attribute("href") ?? "")
                .Where(url => url.StartsWith(prefix, StringComparison.Ordinal))
                .Select(url => url.Substring(prefix.Length))
                .Where(tag => Regex.IsMatch(tag, "\\A[0-9]+\\.[0-9]+\\.[0-9]+\\.0\\z"))
                .Select(tag => { Version v; return Version.TryParse(tag, out v) && v.ToString() == tag ? v : null; })
                .Where(v => v != null).OrderByDescending(v => v).ToArray();
            if (versions.Length == 0) throw new InvalidDataException("GitHub 渠道尚未发布可用版本。");
            return new Uri(GitHubRepository(address), "releases/download/v" + versions[0] + "/latest.xml");
        }

        internal static bool IsGitHubDownloadAddress(Uri uri)
        {
            return uri != null && uri.IsAbsoluteUri && uri.Scheme == "https" && uri.UserInfo.Length == 0 &&
                (uri.Host == "github.com" || uri.Host == "release-assets.githubusercontent.com");
        }

        internal static async Task<HttpResponseMessage> GetResponseAsync(HttpClient client, Uri address, bool github)
        {
            using (var timeout = new CancellationTokenSource(client.Timeout))
            for (int redirects = 0; ; redirects++)
            {
                var response = await client.GetAsync(address, timeout.Token).ConfigureAwait(false);
                int status = (int)response.StatusCode;
                if (status != 301 && status != 302 && status != 303 && status != 307 && status != 308) return response;
                var location = response.Headers.Location;
                response.Dispose();
                Uri next = location == null ? null : new Uri(address, location);
                if (!github || redirects >= 4 || !IsGitHubDownloadAddress(next))
                    throw new InvalidDataException("更新下载重定向到未经允许的地址。");
                address = next;
            }
        }

        private static XElement ReadXml(string xml)
        {
            using (var reader = XmlReader.Create(new StringReader(xml), new XmlReaderSettings
            { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 1024 * 1024 }))
                return XElement.Load(reader);
        }

        internal static UpdateInfo ReadRelease(string manifest, Uri address, string publicKey)
        {
            XElement root = ReadXml(manifest);
            string versionText = (string)root.Attribute("version");
            Version version;
            string file = (string)root.Attribute("file");
            string hash = (string)root.Attribute("sha256") ?? "";
            if (root.Name != "release" || root.HasElements || (string)root.Attribute("product") != "ExcelNavigatorPane" ||
                !Version.TryParse(versionText, out version) || version.Revision != 0 || version.ToString() != versionText ||
                version.Major > 255 || version.Minor > 255 || version.Build > 65535 || version < new Version(1, 0, 1, 0) ||
                file != "releases/" + versionText + "/ExcelNavigator-Setup-" + versionText + ".exe" ||
                !Regex.IsMatch(hash, "\\A[0-9a-f]{64}\\z"))
                throw new InvalidDataException("更新清单格式不正确。");
            byte[] message = Encoding.UTF8.GetBytes("ExcelNavigatorPane\n" + versionText + "\n" + file + "\n" + hash);
            try
            {
                using (RSA rsa = RSA.Create())
                {
                    rsa.FromXmlString(publicKey);
                    if (!rsa.VerifyData(message, Convert.FromBase64String((string)root.Attribute("signature") ?? ""),
                        HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1))
                        throw new InvalidDataException("更新清单签名不正确，已停止下载。");
                }
            }
            catch (Exception ex) when (ex is FormatException || ex is CryptographicException || ex is ArgumentException)
            { throw new InvalidDataException("无法验证更新清单签名。", ex); }
            return new UpdateInfo { Version = version, DownloadUrl = new Uri(address, file), Sha256 = hash };
        }

        internal static async Task<UpdateInfo> CheckAsync()
        {
            XElement settings = ReadSettings();
            Version current = Version.Parse((string)settings.Attribute("version"));
            string channel = (string)settings.Attribute("channel") ?? "oneview";
            Uri address = ParseAddress((string)settings.Attribute("manifestUrl"), channel);
            if (address == null) return new UpdateInfo { Message = "当前版本：" + current + "\n此版本尚未配置更新地址，请在发布时配置。" };
            try
            {
                using (var client = new HttpClient(CreateHttpHandler()))
                {
                    client.Timeout = TimeSpan.FromSeconds(10);
                    client.MaxResponseContentBufferSize = 1024 * 1024;
                    client.DefaultRequestHeaders.CacheControl = new CacheControlHeaderValue { NoCache = true };
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("ExcelNavigatorPane/" + current);
                    Uri manifestAddress = address;
                    if (channel == "github")
                    {
                        using (var listing = await GetResponseAsync(client, address, false).ConfigureAwait(false))
                        {
                            listing.EnsureSuccessStatusCode();
                            manifestAddress = GitHubManifest(await listing.Content.ReadAsStringAsync().ConfigureAwait(false), address);
                        }
                    }
                    using (var response = await GetResponseAsync(client, manifestAddress, channel == "github").ConfigureAwait(false))
                    {
                        if (response.StatusCode == HttpStatusCode.NotFound)
                            return new UpdateInfo { Message = "更新服务尚未提供更新清单。\n当前版本：" + current + "，可继续正常使用。\n请在更新清单发布后重试。" };
                        response.EnsureSuccessStatusCode();
                        var release = ReadRelease(await response.Content.ReadAsStringAsync().ConfigureAwait(false), address,
                            (string)settings.Element("publicKey"));
                        release.GitHub = channel == "github";
                        if (release.GitHub)
                        {
                            var root = GitHubRepository(address);
                            if (manifestAddress != new Uri(root, "releases/download/v" + release.Version + "/latest.xml"))
                                throw new InvalidDataException("GitHub 标签和更新清单版本不一致。");
                            release.DownloadUrl = new Uri(root, "releases/download/v" + release.Version + "/ExcelNavigator-Setup-" + release.Version + ".exe");
                        }
                        if (release.Version <= current) return new UpdateInfo { Message = "当前版本：" + current + "\n暂无更高版本。" };
                        release.Message = "发现新版本：" + release.Version + "（当前 " + current + "）\n是否下载并打开安装程序？安装前会提示你先保存工作，再关闭 Excel 和 WPS 表格。";
                        return release;
                    }
                }
            }
            catch (Exception ex) when (ex is HttpRequestException || ex is TaskCanceledException ||
                ex is XmlException || ex is InvalidDataException)
            { return new UpdateInfo { Message = "无法检查更新，请稍后重试。\n当前版本：" + current + "，可继续正常使用。\n" + ex.Message }; }
        }

        internal static void VerifyInstaller(string path, string expectedHash)
        {
            using (var stream = File.OpenRead(path))
            using (var sha = SHA256.Create())
                if (BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant() != expectedHash)
                    throw new InvalidDataException("安装包校验失败，已丢弃下载文件。");
        }

        internal static async Task DownloadAsync(UpdateInfo release, string destination)
        {
            // Download only; the caller may launch the installer after integrity verification.
            string temporary = destination + "." + Guid.NewGuid().ToString("N") + ".part";
            try
            {
                using (var client = new HttpClient(CreateHttpHandler()))
                {
                    client.Timeout = TimeSpan.FromMinutes(2);
                    client.MaxResponseContentBufferSize = 128 * 1024 * 1024;
                    // Bounded buffering keeps timeout effective for the entire download.
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("ExcelNavigatorPane/" + CurrentVersion);
                    using (var response = await GetResponseAsync(client, release.DownloadUrl, release.GitHub).ConfigureAwait(false))
                    {
                        response.EnsureSuccessStatusCode();
                        byte[] data = await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                        File.WriteAllBytes(temporary, data);
                    }
                }
                VerifyInstaller(temporary, release.Sha256);
                if (File.Exists(destination)) File.Replace(temporary, destination, null);
                else File.Move(temporary, destination);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
