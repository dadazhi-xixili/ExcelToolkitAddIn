using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
namespace ExcelToolkitAddIn
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Layout
    {
        public ThisAddIn addIn;
        public Excel.Application app;
        public Ribbon ribbon;
        public WebView webView;
        //public WebView.ExcelJavaScriptFunctionsWebView excelJavaScriptFunctionsWebView;
        public Sqlite sql = new Sqlite();
        public string appPath = AppDomain.CurrentDomain.BaseDirectory;
        public Udf udf ;
        public Dictionary<WebView.Pane, object> panes = new Dictionary<WebView.Pane, object>();
        public object LoadPane(WebView.Pane pane)
        {
            if (panes.TryGetValue(pane, out object isValue))
            {
                return isValue;
            }
            object newPane;
            switch (pane)
            {
                case WebView.Pane.Query:
                    newPane = new Query();
                    break;
                case WebView.Pane.Name:
                    newPane = new Name();
                    break;
                case WebView.Pane.PowerQuery:
                    newPane = new PowerQuery();
                    break;
                case WebView.Pane.CSharpFunction:
                default:
                    return null;
            }
            panes[pane] = newPane;
            return panes[pane];
        }
    }
}