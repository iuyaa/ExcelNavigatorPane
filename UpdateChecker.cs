using System;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace ExcelNavigatorPane
{
    internal static class UpdateChecker
    {
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
                !uri.AbsolutePath.EndsWith("/ExcelNavigatorPane.vsto", StringComparison.Ordinal))
                throw new InvalidDataException("更新地址必须是固定的 HTTPS ExcelNavigatorPane.vsto 文件地址。");
            return uri;
        }

        internal static Version ReadPublishedVersion(string manifest)
        {
            using (var reader = XmlReader.Create(new StringReader(manifest), new XmlReaderSettings
            { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 1024 * 1024 }))
            {
                XElement root = XElement.Load(reader);
                XNamespace assembly = "urn:schemas-microsoft-com:asm.v1";
                XElement identity = root.Element(assembly + "assemblyIdentity");
                Version version;
                if (root.Name != assembly + "assembly" || identity == null ||
                    (string)identity.Attribute("name") != "ExcelNavigatorPane.vsto" ||
                    !Version.TryParse((string)identity.Attribute("version"), out version) || version.Revision < 0)
                    throw new InvalidDataException("更新地址返回的不是 Excel Navigator 发布清单。");
                // This reads version metadata only. VSTO validates signatures before installing updates.
                return version;
            }
        }

        internal static async Task<string> CheckAsync()
        {
            XElement settings = ReadSettings();
            Version current = Version.Parse((string)settings.Attribute("version"));
            Uri address = ParseAddress((string)settings.Attribute("manifestUrl"));
            if (address == null) return "当前版本：" + current + "\n此版本尚未配置更新地址，请在发布时配置。";
            try
            {
                using (var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false }))
                {
                    client.Timeout = TimeSpan.FromSeconds(10);
                    client.MaxResponseContentBufferSize = 1024 * 1024;
                    client.DefaultRequestHeaders.CacheControl = new CacheControlHeaderValue { NoCache = true };
                    using (var response = await client.GetAsync(address).ConfigureAwait(false))
                    {
                        response.EnsureSuccessStatusCode();
                        Version latest = ReadPublishedVersion(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
                        return latest > current
                            ? "发现新版本：" + latest + "（当前 " + current + "）\n请在方便时保存工作并关闭所有 Excel 窗口，下次启动时由 Office 检查并安装更新。\n本次检查不会关闭 Excel 或安装文件。"
                            : "当前版本：" + current + "\n暂无更高版本。";
                    }
                }
            }
            catch (Exception ex) when (ex is HttpRequestException || ex is TaskCanceledException ||
                ex is XmlException || ex is InvalidDataException)
            {
                return "无法检查更新，请稍后重试。\n当前版本：" + current + "，可继续正常使用。\n" + ex.Message;
            }
        }
    }
}
