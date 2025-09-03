using Microsoft.Office.Tools;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using System;
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
        public System.Windows.Forms.UserControl control;
        public CustomTaskPane controlTaskPane;
        public string htmlPath;
        public Task initTask;
        public Layout layout = Globals.ThisAddIn.layout;
        public Pane pane;
        public Task paneTask;

        public WebView(Pane pane, int size = 1200)
        {
            layout.webView = this;
            control = new System.Windows.Forms.UserControl();
            Dock = DockStyle.Fill;
            controlTaskPane = layout.addIn.CustomTaskPanes.Add(control, "Excel Toolkit");
            control.Controls.Add(this);
            SetSize(size);
            initTask = InitWebView(pane);
            controlTaskPane.Visible = false;
        }

        public new bool Visible
        {
            get => controlTaskPane.Visible;
            set => controlTaskPane.Visible = value;
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

        public void SetSize(int width)
        {
            controlTaskPane.Width = width;
        }

        //public class ExcelJavaScriptFunctionsWebView : WebView2
        //{
        //    public string appPath = AppDomain.CurrentDomain.BaseDirectory;
        //    public ExcelJavaScriptFunctionsWebView(bool isDebug = false)
        //    {
        //        Visible = false;

        //        if (isDebug)
        //        {
        //            Visible = true;
        //            UserControl control = new UserControl();
        //            Dock = DockStyle.Fill;
        //            CustomTaskPane controlTaskPane =  Globals.ThisAddIn.layout.addIn.CustomTaskPanes.Add(control, "ExcelJavaScriptFunctionsWebView");
        //            control.Controls.Add(this);
        //            controlTaskPane.Visible = true;
        //            controlTaskPane.Width = 200;
        //        }
        //        _ = InitWebView();
        //    }
        //    private async Task InitWebView()
        //    {
        //        string htmlFilePath = Path.Combine(appPath, "HTML", "ExcelJavaScriptFunction.html");
        //        string userFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelToolkitWebView2");
        //        var env = await CoreWebView2Environment.CreateAsync(null, userFolder);
        //        await EnsureCoreWebView2Async(env);
        //        string html = File.ReadAllText(htmlFilePath);
        //        NavigateToString(html);  
        //    }
        //}
    }
}