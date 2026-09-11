// Compile with UpdateChecker.cs and packaged UpdateSettings.xml. Read-only live HTTPS check.
using System;
using System.Net;
using ExcelNavigatorPane;

class UpdateTransportCheck
{
    static int Main()
    {
        try
        {
            // Reproduce an older Office/WPS host without changing another process's TLS settings.
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
            var result = UpdateChecker.CheckAsync().GetAwaiter().GetResult();
            if (ServicePointManager.SecurityProtocol != SecurityProtocolType.Tls)
                throw new Exception("Update request changed host-wide TLS settings.");
            if (result.Message.Contains("无法检查更新") || result.Message.Contains("尚未配置更新地址"))
                throw new Exception(result.Message);
            Console.WriteLine("PASS: HTTPS works with legacy host defaults; global TLS unchanged. " + result.Message);
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex); return 1; }
    }
}
