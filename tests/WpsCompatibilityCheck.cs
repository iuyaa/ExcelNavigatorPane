// Offline regression: real hidden test HWNDs and COM interface stand-ins; never attaches to Office.
// csc /r:System.Windows.Forms.dll /out:WpsCompatibilityCheck.exe tests\WpsCompatibilityCheck.cs
// WpsCompatibilityCheck.exe <built ExcelNavigatorPane.dll>
using System;
using System.Drawing;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Runtime.Serialization;
using System.Threading.Tasks;
using System.Windows.Forms;

class WpsCompatibilityCheck
{
    [System.Runtime.InteropServices.DllImport("user32.dll")] static extern int GetWindowLong(IntPtr hwnd,int index);
    [System.Runtime.InteropServices.DllImport("user32.dll")] static extern int SetWindowLong(IntPtr hwnd,int index,int value);
    const BindingFlags F = BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
    class Stub : RealProxy
    {
        readonly Func<IMethodCallMessage, object> handler;
        public Stub(Type type, Func<IMethodCallMessage, object> handler) : base(type) { this.handler = handler; }
        public override IMessage Invoke(IMessage message)
        {
            var call = (IMethodCallMessage)message;
            try { return new ReturnMessage(handler(call), null, 0, call.LogicalCallContext, call); }
            catch (Exception ex) { return new ReturnMessage(ex, call); }
        }
    }
    static object Proxy(Type type, Func<IMethodCallMessage, object> handler) { return new Stub(type,handler).GetTransparentProxy(); }
    static object Call(object target, string name, params object[] args)
    {
        try { return target.GetType().GetMethod(name,F).Invoke(target,args); }
        catch (TargetInvocationException ex) { throw ex.InnerException; }
    }
    static void Set(object target,string field,object value) { target.GetType().GetField(field,F).SetValue(target,value); }
    static void Require(bool ok,string reason) { if(!ok) throw new Exception(reason); }
    static void CheckUpdateUi(Type controlType)
    {
        using(var pane = (Control)Activator.CreateInstance(controlType))
        {
            var link = (LinkLabel)controlType.GetField("_updateLink",F).GetValue(pane);
            var counts = (Label)controlType.GetField("_lblCounts",F).GetValue(pane);
            var dot = (Label)controlType.GetField("_updateDot",F).GetValue(pane);
            var footer = (Panel)counts.Parent;
            string currentVersion = link.Text;
            Require(link.Text.StartsWith("v") && !dot.Visible,"Current version must show without an update dot initially");
            Require(link.Parent==footer && dot.Parent==footer && pane.Controls.Count==1,"Counts and version must share one row, without a second footer");
            Require(!footer.Controls.Cast<Control>().Any(c=>c.Text=="功能说明"),"Feature guide must remain in menu only");
            counts.Text="12 张可见  ·  3 张隐藏";
            var infoType = controlType.Assembly.GetType("ExcelNavigatorPane.UpdateChecker+UpdateInfo");
            object info = Activator.CreateInstance(infoType,true);
            infoType.GetField("Version",F).SetValue(info,new Version(1,0,99,0));
            Set(pane,"_availableUpdate",info);
            Call(pane,"RefreshUpdateStatus");
            Require(link.Text==currentVersion && dot.Visible && dot.ForeColor==Color.Firebrick,
                "New release must add red dot without replacing or recoloring current version");
            Require(link.LinkColor==counts.ForeColor && link.AccessibleDescription.Contains("1.0.99.0"),"Version stays gray; accessible hint names the update");
            foreach(int width in new[]{240,280,340})
            {
                pane.Size = new Size(width,600); pane.PerformLayout(); footer.PerformLayout();
                Require(counts.Right<=link.Left && link.Right<=dot.Left && counts.Width>80,
                    "Counts, version and dot must not overlap at narrow widths");
            }
            Require(controlType.GetField("_updateTimer",F).GetValue(pane)==null,"Uninitialized preview must never start background requests");
            // Feed completed results into the shared cache: exercise the real asynchronous UI path without network or Office.
            var checker = controlType.Assembly.GetType("ExcelNavigatorPane.UpdateChecker");
            var shared = checker.GetField("_sharedCheck",F);
            var checkedAt = checker.GetField("_checkedAt",F);
            object oldTask=shared.GetValue(null), oldTime=checkedAt.GetValue(null);
            try
            {
                IntPtr handle=pane.Handle;
                var contextType=controlType.GetNestedType("PaneSynchronizationContext",BindingFlags.NonPublic);
                Set(pane,"_uiContext",Activator.CreateInstance(contextType,F,null,new object[]{pane},null));
                foreach(bool available in new[]{false,true})
                {
                    object result=Activator.CreateInstance(infoType,true);
                    if(available)
                    {
                        infoType.GetField("Version",F).SetValue(result,new Version(1,0,100,0));
                        infoType.GetField("DownloadUrl",F).SetValue(result,new Uri("https://example.invalid/never-download.exe"));
                    }
                    else infoType.GetField("Message",F).SetValue(result,"offline failure");
                    var completed=typeof(Task).GetMethod("FromResult").MakeGenericMethod(infoType).Invoke(null,new[]{result});
                    shared.SetValue(null,completed); checkedAt.SetValue(null,DateTime.UtcNow);
                    int forms=Application.OpenForms.Count;
                    Call(pane,"CheckUpdates",false);
                    for(int i=0;i<100 && (bool)controlType.GetField("_checkingUpdate",F).GetValue(pane);i++)
                    { Application.DoEvents(); System.Threading.Thread.Sleep(5); }
                    Require(!(bool)controlType.GetField("_checkingUpdate",F).GetValue(pane),"Background UI callback must complete");
                    Require(Application.OpenForms.Count==forms,"Automatic result must never open a dialog");
                    Require(link.Text==currentVersion && dot.Visible && link.AccessibleDescription.Contains(available ? "1.0.100.0" : "1.0.99.0"),
                        "Failure preserves dot; success updates hint, not installed version");
                }
            }
            finally { shared.SetValue(null,oldTask); checkedAt.SetValue(null,oldTime); }
            string statusPreview = Environment.GetEnvironmentVariable("EXCEL_NAVIGATOR_TEST_PREVIEW");
            if(!string.IsNullOrEmpty(statusPreview))
                using(var bitmap=new Bitmap(footer.Width,footer.Height))
                { footer.DrawToBitmap(bitmap,new Rectangle(0,0,footer.Width,footer.Height)); bitmap.Save(Path.ChangeExtension(statusPreview,"status.png")); }
            Set(pane,"_availableUpdate",null); Call(pane,"RefreshUpdateStatus");
            Require(!dot.Visible && link.Text==currentVersion,"No available update must clear dot");
        }
        var dialogType = controlType.GetNestedType("InformationDialog",BindingFlags.NonPublic);
        using(var dialog = (Form)Activator.CreateInstance(dialogType,F,null,new object[]{"更新详情","当前 v1.0.14.0 → 新版 v1.0.15.0",
            "1.0.15.0\r\n\r\n新增功能\r\n• 导航栏更新提示\r\n• 离线功能说明\r\n\r\n问题修复\r\n• 更新说明支持跨版本查看",true},null))
        {
            Require(((Button)dialog.AcceptButton).DialogResult==DialogResult.Cancel,"Enter must not silently start installation");
            dialog.CreateControl(); dialog.PerformLayout();
            var body = dialog.Controls.Cast<Control>().SelectMany(c => c.Controls.Cast<Control>()).OfType<TextBox>().Single();
            Require(body.ReadOnly && body.Multiline,"Release notes must be inert, readable text");
            var actions = dialog.Controls.OfType<FlowLayoutPanel>().Single();
            Require(actions.Controls.OfType<Button>().Count(b => b.DialogResult==DialogResult.OK)==1,"Download requires explicit confirmation");
            dialog.ClientSize = new Size(480,360); dialog.PerformLayout();
            Require(body.Height>120 && body.Width>300,"Resized details must remain readable");
            string preview = Environment.GetEnvironmentVariable("EXCEL_NAVIGATOR_TEST_PREVIEW");
            if(!string.IsNullOrEmpty(preview))
            {
                dialog.ClientSize = new Size(600,500); dialog.PerformLayout();
                // Render only this test window offscreen, with native activation disabled.
                dialog.StartPosition = FormStartPosition.Manual; dialog.Location = new Point(-30000,-30000);
                SetWindowLong(dialog.Handle,-20,GetWindowLong(dialog.Handle,-20) | 0x08000000);
                dialog.Show(); Application.DoEvents();
                using(var bitmap=new Bitmap(dialog.Width,dialog.Height)) { dialog.DrawToBitmap(bitmap,new Rectangle(0,0,dialog.Width,dialog.Height)); bitmap.Save(preview); }
                dialog.Hide();
            }
        }
        Console.WriteLine("PASS: inline version/update state, narrow footer layout, safe update-dialog default and inert notes (offline)");
    }
    static void CheckWorkbookCopy(Type controlType)
    {
        Type bookType=controlType.GetField("_workbook",F).FieldType;
        var getFile=controlType.GetMethod("GetWorkbookFileForClipboard",F);
        var resolve=controlType.GetMethod("ResolveLocalWorkbookPath",F);
        var readable=controlType.GetMethod("CanReadWorkbookFile",F);
        string folder=Path.Combine(Path.GetTempPath(),"ExcelNavigatorOriginalCheck-"+Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        string file=Path.Combine(folder,"Original # 100%.xlsm");
        string actual=file;
        bool saved=false;
        int saves=0, mode=0;
        object book=Proxy(bookType,c => {
            if(c.MethodName=="get_Saved") return saved;
            if(c.MethodName=="get_FullName") { Require(saved,"Must save before resolving actual file"); return actual; }
            if(c.MethodName=="Save") {
                saves++;
                if(mode==2) throw new IOException("Save failed");
                if(mode==0) { File.WriteAllText(file,"latest edits"); saved=true; }
                return null;
            }
            throw new Exception("Actual-file copy must not export, activate or rename source: "+c.MethodName);
        });
        try {
            File.WriteAllText(file,"old contents");
            Require((string)getFile.Invoke(null,new[]{book})==file && saves==1 && File.ReadAllText(file)=="latest edits","Must save and return exact original file, including literal # and percent");
            Require((string)getFile.Invoke(null,new[]{book})==file && saves==1,"Already-saved file needs no extra save or snapshot");
            Require((bool)readable.Invoke(null,new object[]{file}),"Readable original must not warn");
            using(var locked=File.Open(file,FileMode.Open,FileAccess.ReadWrite,FileShare.None))
                Require(!(bool)readable.Invoke(null,new object[]{file}),"Locked original needs warning before paste");
            foreach(int failure in new[]{1,2}) {
                saved=false;mode=failure;
                try { getFile.Invoke(null,new[]{book}); throw new Exception("Canceled/failed save accepted"); }
                catch(TargetInvocationException ex) { Require(ex.InnerException is IOException || ex.InnerException is InvalidOperationException,"Save must fail before clipboard update"); }
            }
            saved=true; actual=Path.Combine(folder,"missing.xlsx");
            try { getFile.Invoke(null,new[]{book}); throw new Exception("Missing original accepted"); }
            catch(TargetInvocationException ex) { Require(ex.InnerException is FileNotFoundException,"Missing original must fail explicitly"); }
            var mappings=new List<KeyValuePair<string,string>> {
                new KeyValuePair<string,string>("https://example.sharepoint.com/Documents/",folder),
                new KeyValuePair<string,string>("https://example.sharepoint.com/Documents/",folder)
            };
            string url="https://example.sharepoint.com/Documents/Original%20%23%20100%25.xlsm";
            Require((string)resolve.Invoke(null,new object[]{url,mappings})==file,"Decode actual filename and merge duplicate mappings");
            Require((string)resolve.Invoke(null,new object[]{"https://example.sharepoint.com/Documents/sub/a+b.xlsx",mappings})==Path.Combine(folder,"sub","a+b.xlsx"),"URL plus sign is literal in a filename");
            foreach(string invalid in new[]{"relative.xlsx","https://other.sharepoint.com/Documents/a.xlsx","https://example.sharepoint.com/Documents-other/a.xlsx","https://example.sharepoint.com/Documents/../escape.xlsx","https://example.sharepoint.com/Documents/%5Cescape.xlsx"}) {
                try { resolve.Invoke(null,new object[]{invalid,mappings}); throw new Exception("Unsafe or unmapped address accepted: "+invalid); }
                catch(TargetInvocationException ex) { Require(ex.InnerException is InvalidOperationException,"Invalid mapping must fail closed"); }
            }
            mappings.Add(new KeyValuePair<string,string>("https://example.sharepoint.com/Documents/",Path.Combine(folder,"other-root")));
            try { resolve.Invoke(null,new object[]{url,mappings}); throw new Exception("Ambiguous roots accepted"); }
            catch(TargetInvocationException ex) { Require(ex.InnerException is InvalidOperationException,"Conflicting sync mappings must not guess"); }
            Require(Directory.GetFiles(folder).Length==1,"Copy must not create extra files");
            foreach(int width in new[]{280,320,400}) {
                var row=new System.Drawing.Rectangle(0,0,width,30);
                var copy=(System.Drawing.Rectangle)controlType.GetMethod("GetWbCopyRect",F).Invoke(null,new object[]{row});
                var close=(System.Drawing.Rectangle)controlType.GetMethod("GetWbCloseRect",F).Invoke(null,new object[]{row});
                Require(copy.Right<=close.Left && row.Contains(copy) && row.Contains(close),"Copy must precede close without overlapping");
            }
        } finally { File.Delete(file); Directory.Delete(folder); }
        Console.WriteLine("PASS: save then copy actual path, canceled/failed save rejection, OneDrive URL mapping and boundaries, no temporary copies; clipboard untouched");
    }
    static void CheckWorksheetRefresh(Type controlType, Type addinType)
    {
        Type appType = controlType.GetField("_app",F).FieldType;
        Type bookType = controlType.GetField("_workbook",F).FieldType;
        Type windowType = controlType.GetField("_window",F).FieldType;
        Type sheetType = addinType.GetMethod("ToggleWorksheetVisibility",F).GetParameters()[0].ParameterType;
        Type sheetsType = bookType.GetInterfaces().Concat(new[]{bookType}).SelectMany(t=>t.GetProperties()).First(p=>p.Name=="Worksheets").PropertyType;
        var sheets = new ArrayList();
        object active = null, book = null;
        var visibilities = Enumerable.Repeat(-1,60).ToArray();
        visibilities[30] = 0;
        Action visibilityChanged = () => {};
        object collection = Proxy(sheetsType,c => c.MethodName=="GetEnumerator" ? sheets.GetEnumerator() : (object)sheets.Count);
        Func<IMethodCallMessage,object> bookHandler = c => c.MethodName=="get_Worksheets" ? collection : c.MethodName=="get_ProtectStructure" ? (object)false : null;
        book = Proxy(bookType,bookHandler);
        for(int i=0;i<60;i++) {
            int index=i;
            object sheet=null;
            sheet=Proxy(sheetType,c => {
                if(c.MethodName=="get_Name" || c.MethodName=="get_CodeName") return "Sheet"+index;
                if(c.MethodName=="get_Visible") return Enum.ToObject(((MethodInfo)c.MethodBase).ReturnType,visibilities[index]);
                if(c.MethodName=="set_Visible") { visibilities[index]=Convert.ToInt32(c.Args[0]); visibilityChanged(); return null; }
                if(c.MethodName=="get_Parent") return book;
                if(c.MethodName=="get_ProtectContents") return false;
                if(c.MethodName=="get_Tab") return null;
                if(c.MethodName=="Activate") { active=sheet; return null; }
                throw new Exception("Unexpected scroll worksheet call: "+c.MethodName);
            });
            sheets.Add(sheet);
        }
        active=sheets[0];
        object app=Proxy(appType,c => c.MethodName=="get_Ready" ? (object)true : null);
        object window=Proxy(windowType,c => c.MethodName=="get_ActiveSheet" ? active : null);
        using(var form=new Form { ClientSize=new System.Drawing.Size(360,700), ShowInTaskbar=false })
        using(var control=(Control)Activator.CreateInstance(controlType)) {
            control.Dock=DockStyle.Fill; form.Controls.Add(control);
            IntPtr handle=form.Handle;
            control.CreateControl(); form.PerformLayout();
            Set(control,"_app",app); Set(control,"_workbook",book); Set(control,"_window",window);
            var list=(ListBox)controlType.GetField("_lbWorksheets",F).GetValue(control);
            handle=list.Handle;
            Action refresh=()=>Call(control,"RefreshWorksheets",true);
            refresh(); Require(list.Items.Count==60,"Scroll fixture must enumerate real refresh path");
            list.TopIndex=25; int top=list.TopIndex;
            Require(top==25,"Scroll fixture must have enough rows");
            refresh(); Require(list.TopIndex==top,"Refresh must preserve scroll even when active sheet is above viewport");
            visibilityChanged=refresh;
            object addin=FormatterServices.GetUninitializedObject(addinType);
            Set(addin,"_windowPanes",Activator.CreateInstance(addinType.GetField("_windowPanes",F).FieldType));
            Set(addin,"_hiddenStates",Activator.CreateInstance(addinType.GetField("_hiddenStates",F).FieldType));
            Set(addin,"<IsWpsHost>k__BackingField",true);
            Set(control,"_addIn",addin);
            Call(addin,"ToggleWorksheetVisibility",sheets[30],control);
            refresh();
            Require(visibilities[30]==-1 && ReferenceEquals(active,sheets[30]),"Unhide must activate target");
            Require(list.TopIndex==top && list.SelectedIndex==30,"Unhide and intermediate refresh must preserve viewport");
            refresh(); Require(list.TopIndex==top,"Repeated event refresh must retain viewport");
            active=sheets[55]; refresh();
            int rows=Math.Max(1,list.ClientSize.Height/list.ItemHeight);
            Require(list.TopIndex==55-rows+1,"Offscreen activation must scroll only enough to reveal target");
            list.TopIndex=25; string anchor=list.Items[list.TopIndex].ToString();
            sheets.RemoveAt(0); refresh();
            Require(list.Items[list.TopIndex].ToString()==anchor,"Removing an earlier sheet must preserve top sheet identity");
            sheets.RemoveAt(list.TopIndex); refresh();
            Require(list.TopIndex>0,"Removing anchor must retain a nearby position");
            active=sheets[0]; Set(control,"_workbook",Proxy(bookType,bookHandler)); refresh();
            Require(list.TopIndex==0,"Switching workbooks must not restore prior workbook scroll");
            var filter=(ToolStripTextBox)controlType.GetField("_txtFilter",F).GetValue(control);
            filter.Text="no matching worksheet"; refresh(); Require(list.Items.Count==0,"Empty filter results must be safe");
            filter.Text=""; refresh(); Require(list.Items.Count==58,"Clearing filter must restore rows");
        }
        using(var dialog=(Form)Activator.CreateInstance(controlType.GetNestedType("WorkbookRenameDialog",F),new object[]{new string('长',100),".xlsx"})) {
            var input=dialog.Controls.OfType<TextBox>().Single();
            Require(dialog.ClientSize.Width>=640 && input.Width>=600,"Rename dialog must show a wider filename field");
            Require(input.Text.Length==100,"Long default name must not be truncated");
            int width=input.Width; dialog.Width+=180; dialog.PerformLayout();
            Require(input.Width==width+180,"Filename field must expand with dialog");
            Require(dialog.AcceptButton.DialogResult==DialogResult.OK && dialog.CancelButton.DialogResult==DialogResult.Cancel,"Rename keyboard confirmation/cancel contract");
        }
        Console.WriteLine("PASS: worksheet scroll preservation, unhide activation, repeated refresh, offscreen activation, removal/filter/book-switch boundaries; wide resizable rename dialog (offline)");
    }
    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            var assembly = Assembly.LoadFrom(args[0]);
            var addinType = assembly.GetType("ExcelNavigatorPane.ThisAddIn");
            var controlType = assembly.GetType("ExcelNavigatorPane.NavigationPaneControl");
            CheckWorkbookCopy(controlType);
            CheckUpdateUi(controlType);
            CheckWorksheetRefresh(controlType,addinType);
            var rename = controlType.GetMethod("GetRenamedWorkbookName", F);
            foreach (var sample in new[] {
                new[] { "Report.xlsx", "Report2026", "Report2026.xlsx" },
                new[] { "Report.xlsx", "Report2026.09", "Report2026.09.xlsx" },
                new[] { "Report.xlsx", "Report.v2", "Report.v2.xlsx" },
                new[] { "Report.xlsx", "New.xlsx", "New.xlsx" },
                new[] { "Report.xlsx", "New.XLSX", "New.xlsx" },
                new[] { "Report.xlsm", "New.v2", "New.v2.xlsm" },
                new[] { "Report.xlsb", "New.xlsb", "New.xlsb" },
                new[] { "Report.xlsx", "New.xlsm", "New.xlsm.xlsx" }
            }) Require((string)rename.Invoke(null,new object[]{sample[0],sample[1]})==sample[2],"Rename must preserve the original extension: " + sample[1]);
            try { rename.Invoke(null,new object[]{"Report.xlsx",".xlsx"}); throw new Exception("Extension-only name accepted"); }
            catch(TargetInvocationException ex) { Require(ex.InnerException is ArgumentException,"Wrong extension-only error"); }
            object addin = FormatterServices.GetUninitializedObject(addinType);
            var dictionary = (IDictionary)Activator.CreateInstance(addinType.GetField("_windowPanes",F).FieldType);
            Set(addin,"_windowPanes",dictionary);
            Set(addin,"_creatingPaneKeys",new HashSet<int>());
            var paneType = dictionary.GetType().GetGenericArguments()[1];
            Type taskPaneType = paneType.GetGenericArguments()[0];
            Type windowType = controlType.GetField("_window",F).FieldType;
            Type workbookType = controlType.GetField("_workbook",F).FieldType;
            Type applicationType = controlType.GetField("_app",F).FieldType;
            Type worksheetType = addinType.GetMethod("ToggleWorksheetVisibility",F).GetParameters()[0].ParameterType;
            var activations = new List<string>();
            bool ready = true;
            var liveWindows = new ArrayList();
            var applicationMembers = applicationType.GetInterfaces().Concat(new[] { applicationType }).SelectMany(t => t.GetProperties()).ToArray();
            Type windowsType = applicationMembers.First(p => p.Name == "Windows").PropertyType;
            object windows = Proxy(windowsType,c => c.MethodName == "GetEnumerator" ? liveWindows.GetEnumerator()
                : c.MethodName == "get_Item" || c.MethodName == "get__Default" ? liveWindows[Convert.ToInt32(c.Args[0])-1] : (object)liveWindows.Count);
            object commandBars = Proxy(applicationMembers.First(p => p.Name == "CommandBars").PropertyType,c => true);
            string bookName = "Original.xlsx";
            object book = Proxy(workbookType,c => {
                if(c.MethodName == "get_Name") return bookName;
                if(c.MethodName == "get_ProtectStructure") return false;
                if(c.MethodName == "Activate") { activations.Add("book"); return null; }
                if(c.MethodName == "Close") { activations.Add("close-book"); return null; }
                if(c.MethodName == "get_Windows") return windows;
                throw new Exception("Unexpected workbook call: " + c.MethodName);
            });
            object books = Proxy(applicationMembers.First(p => p.Name == "Workbooks").PropertyType,
                c => c.MethodName == "GetEnumerator" ? new[] { book }.GetEnumerator() : (object)1);
            object app = Proxy(applicationType,c => {
                if(c.MethodName == "get_Workbooks") return books;
                if(c.MethodName == "get_Ready") return ready;
                if(c.MethodName == "get_CommandBars") return commandBars;
                if(c.MethodName == "get_Windows") return windows;
                // SafeRefresh may query these after navigation. No live workbook means no new pane in this fixture.
                if(c.MethodName == "get_ActiveWindow" || c.MethodName == "get_ActiveWorkbook") return null;
                throw new Exception("Unexpected application call: " + c.MethodName);
            });
            Set(addin,"Application",app);
            Set(addin,"<IsWpsHost>k__BackingField",true);
            using(var frame = new Form())
            using(var otherFrame = new Form())
            using(var firstTab = new Panel())
            using(var secondTab = new Panel())
            using(var control = (Control)Activator.CreateInstance(controlType))
            {
                frame.Controls.Add(firstTab); frame.Controls.Add(secondTab);
                IntPtr hwnd = frame.Handle, tab1 = firstTab.Handle, tab2 = secondTab.Handle;
                Func<IntPtr,object> window = handle => Proxy(windowType,c => {
                    if(c.MethodName == "get_Visible") return true;
                    if(c.MethodName == "get_Hwnd") return handle.ToInt32();
                    if(c.MethodName == "Activate") { activations.Add("window"); return null; }
                    throw new Exception("Unexpected window call: " + c.MethodName);
                });
                object win1 = window(tab1), win2 = window(tab2), winOther = window(otherFrame.Handle);
                int key1 = (int)Call(addin,"GetPaneKey",win1);
                Require(key1 == (int)Call(addin,"GetPaneKey",win2),"WPS tabs must share one pane key");
                Require(key1 != (int)Call(addin,"GetPaneKey",winOther),"Different frames must have different pane keys");
                Set(addin,"<IsWpsHost>k__BackingField",false);
                Require((int)Call(addin,"GetPaneKey",win1) != (int)Call(addin,"GetPaneKey",win2),"Excel keeps independent window keys");
                Set(addin,"<IsWpsHost>k__BackingField",true);
                var resolve = controlType.GetMethod("GetHostFrame",F);
                try { resolve.Invoke(null,new object[]{tab1,(uint)System.Diagnostics.Process.GetCurrentProcess().Id+1}); throw new Exception("Foreign PID accepted"); }
                catch(TargetInvocationException ex) { Require(ex.InnerException is InvalidOperationException,"Wrong PID must fail closed"); }
                Set(control,"_app",app); Set(control,"_window",win1); Set(control,"_workbook",book); Set(control,"_addIn",addin);
                int visible = 0;
                object sheet = Proxy(worksheetType,c => {
                    if(c.MethodName == "get_Parent") return book;
                    if(c.MethodName == "get_Visible") return Enum.ToObject(((MethodInfo)c.MethodBase).ReturnType,visible);
                    if(c.MethodName == "set_Visible") { visible = Convert.ToInt32(c.Args[0]); activations.Add("show"); return null; }
                    if(c.MethodName == "Activate") { activations.Add("sheet"); return null; }
                    throw new Exception("Unexpected worksheet call: " + c.MethodName);
                });
                foreach(bool wps in new[]{true,false}) {
                    Set(addin,"<IsWpsHost>k__BackingField",wps); visible = 0; activations.Clear();
                    Call(addin,"ToggleWorksheetVisibility",sheet,control);
                    Require(visible == -1 && string.Join(",",activations) == "show,window,sheet","Unhide must show then activate in the owning window in both hosts");
                }
                Set(addin,"<IsWpsHost>k__BackingField",true);
                activations.Clear();
                Type itemType = controlType.GetNestedType("WbItem",F);
                object item = Activator.CreateInstance(itemType,true);
                itemType.GetProperty("Workbook").SetValue(item,book,null);
                object otherBook = Proxy(workbookType,c => null);
                control.Size = new System.Drawing.Size(320, 900);
                control.CreateControl(); control.PerformLayout();
                var workbookList = (ListBox)controlType.GetField("_lbWorkbooks",F).GetValue(control);
                var menu = (ContextMenuStrip)controlType.GetField("_workbookContextMenu",F).GetValue(control);
                int menuOpenings = 0;
                menu.Opening += (sender,e) => { menuOpenings++; e.Cancel = true; }; // Exercise dispatch without displaying a menu.
                workbookList.Items.Add(item);
                object otherItem = Activator.CreateInstance(itemType,true);
                itemType.GetProperty("Workbook").SetValue(otherItem,otherBook,null);
                workbookList.Items.Add(otherItem);
                workbookList.SelectedIndex = 1;
                Call(control,"LbWorkbooks_MouseClick",workbookList,new MouseEventArgs(MouseButtons.Right,1,4,workbookList.GetItemRectangle(0).Top+4,0));
                Require(menuOpenings == 1 && ReferenceEquals(controlType.GetField("_workbookMenuTarget",F).GetValue(control),book),"Right-click must bind to clicked workbook without activating it");
                Require(activations.Count == 0,"Context menu must not switch workbooks");
                workbookList.SelectedIndex = 1;
                menu.Items.Cast<ToolStripItem>().Single(i=>i.Text=="关闭").PerformClick();
                Require(string.Join(",",activations)=="close-book","Item command must retain clicked target even if list selection changes");
                activations.Clear();
                Call(control,"LbWorkbooks_MouseClick",workbookList,new MouseEventArgs(MouseButtons.Right,1,4,workbookList.GetItemRectangle(1).Bottom+5,0));
                Require(menuOpenings == 2 && controlType.GetField("_workbookMenuTarget",F).GetValue(control)==null,"Blank area must open a targetless menu");
                foreach(string label in new[]{"保存","重命名...","复制文件","关闭"}) Require(!menu.Items.Cast<ToolStripItem>().Single(i=>i.Text==label).Available,"Blank area must hide target-specific actions");
                foreach(string label in new[]{"新建","打开...","排序","刷新列表","检查更新","功能说明"}) Require(menu.Items.Cast<ToolStripItem>().Single(i=>i.Text==label).Available,"Blank menu missing common action: "+label);
                var headingStrip = (ToolStrip)controlType.GetField("_wbToolStrip",F).GetValue(control);
                headingStrip.Items.Cast<ToolStripItem>().Single(i=>i.AccessibleName=="工作簿操作").PerformClick();
                Require(menuOpenings==3 && ReferenceEquals(controlType.GetField("_workbookMenuTarget",F).GetValue(control),otherBook),"Three-dot entry must reuse the menu with its selected workbook");
                workbookList.Items.Clear();
                Call(control,"LbWorkbooks_MouseClick",workbookList,new MouseEventArgs(MouseButtons.Right,1,4,4,0));
                Require(menuOpenings==4 && controlType.GetField("_workbookMenuTarget",F).GetValue(control)==null,"Empty workbook list must offer common menu");
                Set(control,"_workbook",otherBook); liveWindows.Add(win2);
                // Deliberately leave the cached HWND invalid: WPS must use COM, never XLMAIN foreground logic.
                ((Task)Call(control,"NavigateWorkbookAsync",item)).GetAwaiter().GetResult();
                Require(string.Join(",",activations) == "book,window","WPS workbook navigation must activate the target workbook and window through COM");
                activations.Clear(); ready = false;
                try { ((Task)Call(control,"NavigateWorkbookAsync",item)).GetAwaiter().GetResult(); throw new Exception("Busy WPS accepted navigation"); }
                catch(InvalidOperationException) { }
                Require(activations.Count == 0,"Busy WPS must not call Activate");
                ready = true; liveWindows.Clear();
                // SaveAs can finish while Excel still reports busy: keep the old row until the deferred retry.
                Set(control,"_workbook",null); liveWindows.Add(win1);
                Call(control,"RefreshWorkbooks",false);
                Require(workbookList.Items.Count==1 && workbookList.Items[0].ToString()=="Original.xlsx","Initial workbook name");
                bookName="Renamed.xlsx"; ready=false;
                Call(control,"RefreshWorkbooks",false);
                var retry=(Timer)controlType.GetField("_editRefreshTimer",F).GetValue(control);
                Require(retry!=null && retry.Enabled,"Busy refresh must be retried automatically");
                Call(retry,"OnTick",EventArgs.Empty);
                Require(retry.Enabled && workbookList.Items[0].ToString()=="Original.xlsx","Do not refresh while busy");
                ready=true; Call(retry,"OnTick",EventArgs.Empty);
                Require(!retry.Enabled && workbookList.Items[0].ToString()=="Renamed.xlsx","Retry must replace stale name after save");
                Set(addin,"_refreshQueued",true); Set(addin,"_queuedFullRefresh",false);
                Call(addin,"Application_WorkbookAfterSave",book,false);
                Require(!(bool)addinType.GetField("_queuedFullRefresh",F).GetValue(addin),"Failed save must not enqueue a rename refresh");
                Call(addin,"Application_WorkbookAfterSave",book,true);
                Require((bool)addinType.GetField("_queuedFullRefresh",F).GetValue(addin),"Successful save must enqueue all workbook lists");
                Set(addin,"_refreshQueued",false); Set(addin,"_queuedFullRefresh",false);
                Set(control,"_workbook",book); liveWindows.Clear();
                int created = 0, removed = 0;
                object stalePane = Proxy(taskPaneType,c => c.MethodName == "get_Window" ? win1 : c.MethodName == "get_Visible" ? (object)true : null);
                Set(addin,"CustomTaskPanes",Proxy(addinType.GetField("CustomTaskPanes",F).FieldType,c => {
                    if(c.MethodName == "Add") {
                        created++; Require(c.ArgCount == 2,"WPS pane must not bind to a document");
                        if(created == 1) Call(addin,"EnsurePaneForWindow",win2,otherBook);
                        return stalePane;
                    }
                    if(c.MethodName == "Remove") { removed++; return null; }
                    throw new Exception("Unexpected pane collection call: " + c.MethodName);
                }));
                Call(addin,"EnsurePaneForWindow",win1,book);
                Call(addin,"EnsurePaneForWindow",win2,otherBook);
                Call(addin,"EnsurePaneForWindow",win1,book);
                Require(created == 1 && dictionary.Count == 1,"Opening/switching tabs must reuse one actual pane");
                var sharedControl = (Control)paneType.GetField("Item2").GetValue(dictionary[key1]);
                Require((bool)Call(sharedControl,"IsBoundToWorkbook",book),"Shared pane must bind to the active workbook");
                liveWindows.Add(win2); // original tab closed; remaining tab shares the same frame
                Call(addin,"CleanupClosedPanes");
                Require(dictionary.Count == 1 && !sharedControl.IsDisposed && removed == 0,"Closing one WPS tab must retain the shared pane");
                liveWindows.Clear();
                Call(addin,"CleanupClosedPanes");
                Require(dictionary.Count == 0 && sharedControl.IsDisposed && removed == 1,"Closing the last tab must remove/dispose the pane even when Pane.Window remains non-null");
            }
            Console.WriteLine("PASS: save success/failure and deferred rename refresh, shared/distinct frame keys, Excel isolation, foreign PID rejection, unhide activation order, COM workbook switch, busy guard, tab-close and last-close cleanup (offline)");
            return 0;
        }
        catch(Exception ex) { Console.WriteLine(ex); return 1; }
    }
}
