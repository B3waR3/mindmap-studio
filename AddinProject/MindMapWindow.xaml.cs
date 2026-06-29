using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Web.WebView2.Core;

namespace MindMapStudio
{
    public partial class MindMapWindow : Window
    {
        private const string AddinUrl =
            "https://B3waR3.github.io/mindmap-studio/taskpane/taskpane.html";

        private readonly Connect _connect;

        public MindMapWindow(Connect connect)
        {
            _connect = connect;
            InitializeComponent();
            Loaded += Window_Loaded;
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            await InitWebView();
        }

        private async Task InitWebView()
        {
            try
            {
                // Store WebView2 user data in %LOCALAPPDATA%\MindMapStudio\WebView2
                var userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MindMapStudio", "WebView2");

                var env = await CoreWebView2Environment.CreateAsync(
                    userDataFolder: userDataFolder);

                await webView.EnsureCoreWebView2Async(env);

                // Listen for messages from the web page (e.g. "insert into OneNote")
                webView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;

                // Navigate to the hosted app
                webView.Source = new Uri(AddinUrl);

                // Hide loading overlay once the page finishes loading
                webView.NavigationCompleted += (s, args) =>
                {
                    Dispatcher.Invoke(() => loadingOverlay.Visibility = Visibility.Collapsed);
                };
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "WebView2 Runtime not found.\n\n" +
                    "Please install it from:\nhttps://go.microsoft.com/fwlink/p/?LinkId=2124703\n\n" +
                    $"Error: {ex.Message}",
                    "Mind Map Studio — Setup Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                Close();
            }
        }

        // ── Handle messages from the JavaScript side ────────────────────────────

        private void OnWebMessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                using var doc = JsonDocument.Parse(e.WebMessageAsJson);
                var root = doc.RootElement;

                if (!root.TryGetProperty("type", out var typeProp)) return;
                var type = typeProp.GetString();

                if (type == "insert" && root.TryGetProperty("data", out var dataProp))
                {
                    var base64 = dataProp.GetString();
                    // Strip the data URL header ("data:image/png;base64,")
                    var comma = base64.IndexOf(',');
                    if (comma >= 0) base64 = base64.Substring(comma + 1);

                    Dispatcher.Invoke(() => InsertIntoOneNote(base64));
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() =>
                    MessageBox.Show("Error reading message: " + ex.Message));
            }
        }

        // ── Insert PNG into the current OneNote page via COM API ────────────────

        private void InsertIntoOneNote(string base64Png)
        {
            try
            {
                var app = _connect.OneNoteApp;
                if (app == null)
                {
                    MessageBox.Show("OneNote connection lost. Please restart OneNote.");
                    return;
                }

                // Get the ID of the currently visible page
                string pageId = GetCurrentPageId(app);
                if (string.IsNullOrEmpty(pageId))
                {
                    MessageBox.Show("Could not detect the active OneNote page.\nMake sure a page is open in OneNote.");
                    return;
                }

                // Build minimal page-update XML containing just the new image
                var xml = BuildImageXml(pageId, base64Png);
                app.UpdatePageContent(xml, DateTime.MinValue, 43 /*xs2013*/, false);

                // Notify the web page that insertion succeeded
                webView.CoreWebView2.PostWebMessageAsString("{\"type\":\"insertOk\"}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Could not insert into OneNote:\n" + ex.Message,
                    "Mind Map Studio",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);

                webView.CoreWebView2.PostWebMessageAsString("{\"type\":\"insertError\"}");
            }
        }

        private static string GetCurrentPageId(dynamic app)
        {
            try
            {
                // OneNote COM API: Application.Windows.CurrentWindow.CurrentPageId
                return (string)app.Windows.CurrentWindow.CurrentPageId;
            }
            catch
            {
                return null;
            }
        }

        private static string BuildImageXml(string pageId, string base64Png)
        {
            // Minimal OneNote XML that appends a new image outline to the page.
            // UpdatePageContent with only new elements merges them into the existing page.
            return
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<one:Page xmlns:one=\"http://schemas.microsoft.com/office/onenote/2013/onenote\"" +
                $" ID=\"{pageId}\">" +
                "<one:Outline>" +
                "<one:OEChildren>" +
                "<one:OE>" +
                "<one:Image format=\"png\">" +
                $"<one:Data>{base64Png}</one:Data>" +
                "</one:Image>" +
                "</one:OE>" +
                "</one:OEChildren>" +
                "</one:Outline>" +
                "</one:Page>";
        }

        // ── Keep the window closeable without killing the add-in ────────────────

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Just hide instead of destroying so it reopens quickly
            e.Cancel = true;
            Hide();
        }
    }
}
