using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace ExcelToolkitAddIn
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        public Layout layout;
        public string[] level1;
        public string level1Active;
        public Office.IRibbonUI ribbon;
        public WebView webView;
        public Xml xml;
        public Dictionary<string, bool> isChecks = new Dictionary<string, bool>();

        #region IRibbonExtensibility 成员
        /// <summary>
        /// 获取Ribbon XML
        /// 在调用ribbon.Invalidate时会重新获取
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public string GetCustomUI(string id)
        {
            _ = id;
            InitXml();
            return xml.ToXml();
        }

        /// <summary>
        /// 初始化xml类
        /// 只执行一次
        /// </summary>
        private void InitXml()
        {
            if (layout != null) return;
            layout = Globals.ThisAddIn.layout;
            layout.ribbon = this;
            Query query = new Query();
            layout.panes[WebView.Pane.Query] = query;
            level1 = query.GetLevel1();

            Xml.IControl buttonPQ = new Xml.Button("PQ查询", "PQ查询", "PowerQueryCheckClick", "large", "ViewDocumentMap");
            //Xml.IControl buttonName = new Xml.Button("名称管理", "名称管理", "NameClick", "large", "NameDefine");
            Xml.Group groupName = new Xml.Group("工具", "工具", buttonPQ);

            Xml.IControl[] buttonsQuery = level1.Select(item => (Xml.IControl)new Xml.Button(item, item, "QueryClick")).ToArray();
            Xml.IControl splitButtonQuery = new Xml.SplitButton("函数分类", "分类", "SplitButtonQueryClick", "ShapeSheetShowFormulas", buttonsQuery);
            Xml.Group groupQuery = new Xml.Group("函数查询", "函数查询", splitButtonQuery);

            Xml.Tab tabToolkit = new Xml.Tab("Toolkit", "Toolkit", groupName, groupQuery);
            this.xml = new Xml(tabToolkit);
        }
        #endregion

        #region 功能区回调

        #region CheckedBox交互

        public bool GetChecked(Office.IRibbonControl control)
        {
            if (isChecks.TryGetValue(control.Id, out bool isValue))
            {
                return isChecks[control.Id];
            }
            isChecks.Add(control.Id, isValue);
            return isValue;
        }
        public void CheckBoxClick(Office.IRibbonControl control, bool pressed)
        {
            isChecks[control.Id] = pressed;
        }

        #endregion

        #region WebView加载
        private void LoadWebView(WebView.Pane pane)
        {
            if (webView != null) return;
            webView = new WebView(pane);
            layout.webView = webView;
        }

        #endregion

        #region 函数工具窗格交互
        /// <summary>
        /// 函数分类Level1按钮点击事件
        /// </summary>
        /// <param name="control">所点击的按钮</param>
        public void QueryClick(Office.IRibbonControl control)
        {
            const WebView.Pane pane = WebView.Pane.Query;
            LoadWebView(pane);
            if (webView.pane != pane) webView.LoadHtml(pane);
            if (control.Id == level1Active)
            {
                webView.Visible = !webView.Visible;
            }
            else
            {
                _ = webView.RunJavaScript($"InitLevel2('{control.Id}')");
                level1Active = control.Id;
                webView.Visible = true;
            }
        }
        /// <summary>
        /// 显示函数工具窗格
        /// </summary>
        /// <param name="control"></param>
        public void SplitButtonQueryClick(Office.IRibbonControl control)
        {
            _ = control;
            const WebView.Pane pane = WebView.Pane.Query;
            LoadWebView(pane);
            if (level1Active == null)
            {
                level1Active = level1[0];
                webView.Visible = true;
                _ = webView.RunJavaScript($"InitLevel2('{level1Active}')");
            }
            else
            {
                if (webView.pane == pane)
                {
                    webView.Visible = !webView.Visible;
                }
                else
                {
                    webView.LoadHtml(pane);
                    _ = webView.RunJavaScript($"InitLevel2('{level1Active}')");
                }

            }
        }


        #endregion

        #region 名称窗格交互
        /// <summary>
        /// 名称窗格交互
        /// </summary>
        /// <param name="control">按钮本身</param>
        public void NameClick(Office.IRibbonControl control)
        {
            _ = control;
            const WebView.Pane pane = WebView.Pane.Name;
            LoadWebView(pane);
            if (webView.pane != pane)
            {
                webView.LoadHtml(pane);
                webView.Visible = true;
            }
            else
            {
                webView.Visible = !webView.Visible;
            }
        }

        #endregion

        #region PowerQuery工具窗格
        public void PowerQueryCheckClick(Office.IRibbonControl control)
        {
            _ = control;
            const WebView.Pane pane = WebView.Pane.PowerQuery;
            LoadWebView(pane);
            if (webView.pane != pane)
            {
                webView.LoadHtml(pane);
                webView.SetSize(800);
                webView.controlTaskPane.Visible = true;
            }
            else
            {
                webView.SetSize(800);
                webView.Visible = !webView.Visible;
            }
        }
        #endregion

        #endregion

        #region ribbon UI控制

        /// <summary>
        /// 刷新ribbon
        /// </summary>
        public void RefreshRibbon()
        {
            ribbon.Invalidate();
        }
        public void Ribbon_Load(Office.IRibbonUI ui)
        {
            ribbon = ui;
        }

        #endregion

    }
}