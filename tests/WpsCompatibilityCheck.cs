// Offline regression: real hidden test HWNDs and COM interface stand-ins; never attaches to Office.
// csc /r:System.Windows.Forms.dll /out:WpsCompatibilityCheck.exe tests\WpsCompatibilityCheck.cs
// WpsCompatibilityCheck.exe <built ExcelNavigatorPane.dll>
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Runtime.Serialization;
using System.Threading.Tasks;
using System.Windows.Forms;

class WpsCompatibilityCheck
{
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
    [STAThread]
    static int Main(string[] args)
    {
        try
        {
            var assembly = Assembly.LoadFrom(args[0]);
            var addinType = assembly.GetType("ExcelNavigatorPane.ThisAddIn");
            var controlType = assembly.GetType("ExcelNavigatorPane.NavigationPaneControl");
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
                foreach(string label in new[]{"保存","重命名...","关闭"}) Require(!menu.Items.Cast<ToolStripItem>().Single(i=>i.Text==label).Available,"Blank area must hide target-specific actions");
                foreach(string label in new[]{"新建","打开...","排序","刷新列表","检查更新","关于导航栏"}) Require(menu.Items.Cast<ToolStripItem>().Single(i=>i.Text==label).Available,"Blank menu missing common action: "+label);
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
