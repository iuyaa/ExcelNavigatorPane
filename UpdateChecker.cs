using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Security.Cryptography;
using System.Security.Authentication;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
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
        }

        internal static XElement ReadSettings()
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("ExcelNavigatorPane.UpdateSettings.xml"))
                return XElement.Load(stream);
        }

        internal static Version CurrentVersion => Version.Parse((string)ReadSettings().Attribute("version"));

        internal static Uri ParseAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address)) return null;
            Uri uri;
            if (!Uri.TryCreate(address, UriKind.Absolute, out uri) || uri.Scheme != Uri.UriSchemeHttps ||
                uri.UserInfo.Length != 0 || uri.Query.Length != 0 || uri.Fragment.Length != 0 ||
                !uri.AbsolutePath.EndsWith("/latest.xml", StringComparison.Ordinal))
                throw new InvalidDataException("更新地址必须是固定的 HTTPS latest.xml 文件地址。");
            return uri;
        }

        internal static UpdateInfo ReadRelease(string manifest, Uri address, string publicKey)
        {
            XElement root;
            using (var reader = XmlReader.Create(new StringReader(manifest), new XmlReaderSettings
            { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 1024 * 1024 }))
                root = XElement.Load(reader);
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
            Uri address = ParseAddress((string)settings.Attribute("manifestUrl"));
            if (address == null) return new UpdateInfo { Message = "当前版本：" + current + "\n此版本尚未配置更新地址，请在发布时配置。" };
            try
            {
                using (var client = new HttpClient(CreateHttpHandler()))
                {
                    client.Timeout = TimeSpan.FromSeconds(10);
                    client.MaxResponseContentBufferSize = 1024 * 1024;
                    client.DefaultRequestHeaders.CacheControl = new CacheControlHeaderValue { NoCache = true };
                    using (var response = await client.GetAsync(address).ConfigureAwait(false))
                    {
                        if (response.StatusCode == HttpStatusCode.NotFound)
                            return new UpdateInfo { Message = "更新服务尚未提供更新清单。\n当前版本：" + current + "，可继续正常使用。\n请在更新清单发布后重试。" };
                        response.EnsureSuccessStatusCode();
                        var release = ReadRelease(await response.Content.ReadAsStringAsync().ConfigureAwait(false), address,
                            (string)settings.Element("publicKey"));
                        if (release.Version <= current) return new UpdateInfo { Message = "当前版本：" + current + "\n暂无更高版本。" };
                        release.Message = "发现新版本：" + release.Version + "（当前 " + current + "）\n是否下载安装包？下载后请保存工作、关闭 Excel 和 WPS 表格，再运行安装包升级。";
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
            // Never execute network content; the user runs the verified installer after saving Excel.
            string temporary = destination + "." + Guid.NewGuid().ToString("N") + ".part";
            try
            {
                using (var client = new HttpClient(CreateHttpHandler()))
                {
                    client.Timeout = TimeSpan.FromMinutes(2);
                    client.MaxResponseContentBufferSize = 128 * 1024 * 1024;
                    // Bounded buffering keeps timeout effective for the entire download.
                    byte[] data = await client.GetByteArrayAsync(release.DownloadUrl).ConfigureAwait(false);
                    File.WriteAllBytes(temporary, data);
                }
                VerifyInstaller(temporary, release.Sha256);
                if (File.Exists(destination)) File.Replace(temporary, destination, null);
                else File.Move(temporary, destination);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
