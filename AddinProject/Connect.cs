using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace MindMapStudio
{
    /// <summary>
    /// Static bootstrapper — runs before ANY instance code.
    /// Redirects .NET assembly resolution to the DLL's own folder so that
    /// System.Text.Json, WebView2, etc. are found even though COM loads us
    /// from OneNote's working directory.
    /// </summary>
    internal static class Bootstrapper
    {
        // Safe directory resolution — Assembly.Location can be empty under COM activation
        private static readonly string InstallDir = GetInstallDir();
        private static readonly string LogFile    = Path.Combine(InstallDir, "startup.log");

        private static string GetInstallDir()
        {
            try
            {
                string loc = typeof(Bootstrapper).Assembly.Location;
                if (!string.IsNullOrEmpty(loc))
                {
                    string dir = Path.GetDirectoryName(loc);
                    if (!string.IsNullOrEmpty(dir)) return dir;
                }
            }
            catch { }
            // Fallback: well-known install path
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MindMapStudio");
        }

        static Bootstrapper()
        {
            // Costura.Fody handles assembly resolution for embedded deps.
            // We just log startup here.
            Log("Bootstrapper init. InstallDir=" + InstallDir);
        }

        internal static void Init() { /* ensures static ctor fires */ }

        internal static void Log(string msg)
        {
            try
            {
                File.AppendAllText(LogFile,
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + "\r\n");
            }
            catch { }
        }
    }

    // ── COM interfaces defined manually (no PIA required) ──────────────────────

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup      = 1,
        ext_cm_External     = 2,
        ext_cm_CommandLine  = 3
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed   = 1
    }

    // Note: [ComImport] is intentionally NOT used on these interfaces.
    // [ComImport] means "import from COM", but here we are IMPLEMENTING the interfaces
    // in .NET and exposing them TO COM via the CCW. [ComImport] would confuse the CCW.

    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    public interface IDTExtensibility2
    {
        void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            ext_ConnectMode ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            ref Array custom);

        void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom);
        void OnAddInsUpdate(ref Array custom);
        void OnStartupComplete(ref Array custom);
        void OnBeginShutdown(ref Array custom);
    }

    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    public interface IRibbonExtensibility
    {
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    // IRibbonControl IS a COM object passed to us, so [ComImport] is correct here.
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    public interface IRibbonControl
    {
        string Id      { [return: MarshalAs(UnmanagedType.BStr)] get; }
        object Context { [return: MarshalAs(UnmanagedType.IDispatch)] get; }
        string Tag     { [return: MarshalAs(UnmanagedType.BStr)] get; }
    }

    // ── Main COM add-in class ───────────────────────────────────────────────────

    [ComVisible(true)]
    [Guid("C1A2B3D4-E5F6-7A8B-9C0D-E1F2A3B4C5D6")]
    [ProgId("MindMapStudio.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        internal dynamic OneNoteApp { get; private set; }
        private MindMapWindow _window;

        // Called by OneNote when the add-in is loaded
        public void OnConnection(
            object Application,
            ext_ConnectMode ConnectMode,
            object AddInInst,
            ref Array custom)
        {
            Bootstrapper.Init(); // must be first — sets up assembly resolution
            Bootstrapper.Log("OnConnection called. ConnectMode=" + ConnectMode);
            OneNoteApp = Application;
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            _window?.Close();
            _window = null;
            OneNoteApp = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        // Returns ribbon XML definition (embedded resource)
        public string GetCustomUI(string RibbonID)
        {
            try
            {
                Bootstrapper.Log("GetCustomUI called for: " + RibbonID);
                var asm = Assembly.GetExecutingAssembly();
                using (var stream = asm.GetManifestResourceStream("MindMapStudio.ribbon.xml"))
                using (var reader = new StreamReader(stream))
                {
                    var xml = reader.ReadToEnd();
                    Bootstrapper.Log("GetCustomUI returning " + xml.Length + " chars");
                    return xml;
                }
            }
            catch (Exception ex)
            {
                Bootstrapper.Log("GetCustomUI ERROR: " + ex);
                return "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" />";
            }
        }

        // Called when the ribbon button is clicked
        public void OpenMindMapCallback(IRibbonControl control)
        {
            try
            {
                Bootstrapper.Log("OpenMindMapCallback clicked");
                if (_window == null || !_window.IsLoaded)
                {
                    _window = new MindMapWindow(this);
                    _window.Show();
                }
                else
                {
                    _window.Activate();
                    if (!_window.IsVisible)
                        _window.Show();
                }
            }
            catch (Exception ex)
            {
                Bootstrapper.Log("OpenMindMapCallback ERROR: " + ex);
                System.Windows.MessageBox.Show(
                    "Error opening Mind Map Studio:\n" + ex.Message,
                    "Mind Map Studio");
            }
        }
    }
}
