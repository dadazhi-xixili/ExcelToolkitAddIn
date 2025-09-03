using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToolkitAddIn
{
    public class Udf
    {
        
        public Layout layout = Globals.ThisAddIn.layout;
        public string xllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DNA", "Function-AddIn64-packed.xll");
        public Udf()
        {
            this.Close();
            if(layout.app.RegisterXLL(xllPath))
            {
                System.Diagnostics.Debug.WriteLine("加载成功");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("加载失败");
            }
        }

        public void Close()
        {
            foreach (Excel.AddIn ai in layout.app.AddIns)
            {
                if (!ai.FullName.Equals(xllPath, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                ai.Installed = false;
                System.Diagnostics.Debug.WriteLine("卸载成功");
                break;
            }
            System.Diagnostics.Debug.WriteLine("卸载失败");
        }
    }
}
