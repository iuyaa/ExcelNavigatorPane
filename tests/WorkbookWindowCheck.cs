// csc /r:System.Windows.Forms.dll tests\WorkbookWindowCheck.cs
// WorkbookWindowCheck.exe <ExcelNavigatorPane.dll>
// Creates only its own native test window; does not attach to Excel or suspend Windows.
using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

class WorkbookWindowCheck
{
    delegate IntPtr WindowProc(IntPtr hwnd, uint message, IntPtr wparam, IntPtr lparam);
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct WindowClass
    {
        public uint Style;
        public WindowProc Procedure;
        public int ClassExtra, WindowExtra;
        public IntPtr Instance, Icon, Cursor, Background;
        public string Menu, Name;
    }
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern ushort RegisterClass(ref WindowClass value);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr CreateWindowEx(uint exStyle, string className, string title, uint style, int x, int y, int w, int h, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);
    [DllImport("user32.dll")] static extern IntPtr DefWindowProc(IntPtr h, uint m, IntPtr w, IntPtr l);
    [DllImport("user32.dll")] static extern bool DestroyWindow(IntPtr h);
    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr h, int command);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    static WindowProc procedure = DefWindowProc;

    [STAThread]
    static int Main(string[] args)
    {
        IntPtr window = IntPtr.Zero;
        try
        {
            var type = Assembly.LoadFrom(args[0]).GetType("ExcelNavigatorPane.NavigationPaneControl");
            var method = type.GetMethod("ForegroundExcelWindow", BindingFlags.Static | BindingFlags.NonPublic);
            var wc = new WindowClass { Name = "XLMAIN", Procedure = procedure, Instance = Marshal.GetHINSTANCE(typeof(WorkbookWindowCheck).Module) };
            if (RegisterClass(ref wc) == 0) throw new Exception("Cannot register test window");
            window = CreateWindowEx(0, wc.Name, "ExcelPane window restoration check", 0x00CF0000, 100, 100, 420, 200, IntPtr.Zero, IntPtr.Zero, wc.Instance, IntPtr.Zero);
            if (window == IntPtr.Zero) throw new Exception("Cannot create test window");
            uint pid = (uint)Process.GetCurrentProcess().Id;
            ShowWindow(window, 6);
            if (!IsIconic(window)) throw new Exception("Test window must start minimized");
            bool rejected = false;
            try { method.Invoke(null, new object[] { window, pid + 1 }); }
            catch (TargetInvocationException ex) { rejected = ex.InnerException is InvalidOperationException; }
            if (!rejected || !IsIconic(window)) throw new Exception("Wrong PID must be rejected without restoring the window");
            for (int repeat = 0; repeat < 2; repeat++)
            {
                var task = Task.Run(() => method.Invoke(null, new object[] { window, pid }));
                var timeout = Stopwatch.StartNew();
                while (!task.IsCompleted && timeout.ElapsedMilliseconds < 5000) { Application.DoEvents(); Thread.Sleep(5); }
                if (!task.IsCompleted) throw new Exception("Navigation timed out");
                task.GetAwaiter().GetResult();
                if (IsIconic(window) || GetForegroundWindow() != window) throw new Exception("Navigation returned without restoring and foregrounding the window");
            }
            DestroyWindow(window);
            rejected = false;
            try { method.Invoke(null, new object[] { window, pid }); }
            catch (TargetInvocationException ex) { rejected = ex.InnerException is InvalidOperationException; }
            window = IntPtr.Zero;
            if (!rejected) throw new Exception("Destroyed HWND must be rejected");
            Console.WriteLine("PASS: minimized window restored, foreground verified, repeated activation, wrong PID and stale HWND rejected");
            return 0;
        }
        catch (Exception ex) { Console.WriteLine(ex); return 1; }
        finally { if (window != IntPtr.Zero) DestroyWindow(window); }
    }
}
