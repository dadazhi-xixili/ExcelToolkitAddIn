using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelToolkitAddIn
{
    public partial class ThisAddIn
    {
        public Layout layout = new Layout();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            layout.addIn = this;
            layout.app = Application;
            layout.udf = new Udf();
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            layout.sql.Close();
            layout.udf.Close();
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
