// csc /out:UpdateCheck.exe tests\UpdateCheck.cs
// UpdateCheck.exe <built ExcelNavigatorPane.dll> <real published .vsto file>
using System;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;

class UpdateCheck
{
    static Type checker;
    static object Call(string name, params object[] args)
    {
        try { return checker.GetMethod(name, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, args); }
        catch (TargetInvocationException ex) { throw ex.InnerException; }
    }
    static void Reject(string method, string value)
    {
        bool rejected = false;
        try { Call(method, value); }
        catch (InvalidDataException) { rejected = true; }
        catch (System.Xml.XmlException) { rejected = true; }
        if (!rejected) throw new Exception("Invalid input accepted: " + value);
    }
    static int Main(string[] args)
    {
        try
        {
            checker = Assembly.LoadFrom(args[0]).GetType("ExcelNavigatorPane.UpdateChecker");
            if (Call("ParseAddress", "") != null) throw new Exception("Empty configuration must be disabled");
            Call("ParseAddress", "https://updates.example.invalid/excel/ExcelNavigatorPane.vsto");
            foreach (string address in new[] { "http://host/ExcelNavigatorPane.vsto", "https://user:secret@host/ExcelNavigatorPane.vsto",
                "https://host/", "https://host/ExcelNavigatorPane.vsto?token=temporary", "file:///C:/update.vsto" }) Reject("ParseAddress", address);
            string manifest = File.ReadAllText(args[1]);
            var published = (Version)Call("ReadPublishedVersion", manifest);
            string original = "version=\"" + published + "\"";
            if ((Version)Call("ReadPublishedVersion", manifest.Replace(original, "version=\"9.0.0.0\"")) <= published)
                throw new Exception("Newer version comparison failed");
            Reject("ReadPublishedVersion", "<html><body>Website, not an update</body></html>");
            Reject("ReadPublishedVersion", manifest.Replace("ExcelNavigatorPane.vsto", "AnotherAddIn.vsto"));
            Reject("ReadPublishedVersion", manifest.Replace(original, "version=\"invalid\""));
            Reject("ReadPublishedVersion", "<!DOCTYPE assembly [<!ENTITY x SYSTEM 'file:///secret'>]><assembly>&x;</assembly>");
            Reject("ReadPublishedVersion", new string('x', 1024 * 1024 + 1));
            string result = ((Task<string>)Call("CheckAsync")).GetAwaiter().GetResult();
            if (!result.Contains("尚未配置更新地址")) throw new Exception("Run this check against the unconfigured build");
            Console.WriteLine("PASS: real VSTO version, newer version, disabled configuration, HTTPS address validation, wrong product/HTML/DTD/oversized response rejection");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex); return 1; }
    }
}
