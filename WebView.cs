using Microsoft.Office.Tools;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelToolkitAddIn
{
    public class WebView : WebView2
    {
        /// <summary>
        /// 窗格
        /// Name: 名称窗格
        /// Query: 函数工具窗格
        /// PowerQuery: PowerQuery工具窗格
        /// </summary>
        public enum Pane
        {
            Query,
            Name,
            PowerQuery,
            CSharpFunction
        }

        public string appPath = AppDomain.CurrentDomain.BaseDirectory;
        public UserControl control;
        public string htmlPath;
        public Task initTask;
        public Layout layout = Globals.ThisAddIn.layout;
        public Pane pane;
        public Task paneTask;
        public Form webViewForm;

        public WebView(Pane pane, int size = 1200)
        {
            layout.webView = this;
            control = new UserControl();
            control.Dock = DockStyle.Fill;
            webViewForm = new Form
            {
                Text = @"Excel Toolkit",
                Width = size,
                Height = 800
            };
            webViewForm.Icon = new Icon(Path.Combine(appPath, "Resource", "ExcelToolkit.ico"));
            webViewForm.Controls.Add(control);
            control.Controls.Add(this);
            this.Dock = DockStyle.Fill;
            initTask = InitWebView(pane);
            webViewForm.FormClosing += WebViewFormCloseClick;
        }
        public new bool Visible
        {
            get => webViewForm.Visible;
            set => webViewForm.Visible = value;
        }
        private async Task InitWebView(Pane pane)
        {
            this.pane = pane;
            string userFilePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            userFilePath = Path.Combine(userFilePath, "Excel Toolkit", "webView");
            CoreWebView2Environment env = await CoreWebView2Environment.CreateAsync(null, userFilePath);
            await EnsureCoreWebView2Async(env);
            paneTask = LoadHtml(pane);
        }
        public Task LoadHtml(Pane pane)
        {
            CoreWebView2.AddHostObjectToScript("Layout", layout.LoadPane(pane));
            this.pane = pane;
            htmlPath = Path.Combine(appPath, "HTML", pane + ".html");
            string html = File.ReadAllText(htmlPath);

            NavigateToString(html);
            return Task.CompletedTask;
        }
        public async Task RunJavaScript(string jsCode)
        {
            await initTask;
            await paneTask;
            await CoreWebView2.ExecuteScriptAsync($"layout.{jsCode}");
        }
        private void WebViewFormCloseClick(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            webViewForm.Hide();
        }
    }
}