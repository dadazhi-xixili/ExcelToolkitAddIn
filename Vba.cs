//using System.Linq;
//using Microsoft.Office.Interop.Excel;
//using Microsoft.Vbe.Interop;
//namespace ExcelToolkitAddIn
//{
//    public class Vba
//    {
//        public static void AddModule(Workbook book, string moduleName, string vbaCode)
//        {
//            Microsoft.Vbe.Interop.VBProject vbProject = book.VBProject;
//            VBComponent vbModule = vbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
//            vbModule.Name = moduleName;
//            vbModule.CodeModule.AddFromString(vbaCode);
//        }
//        public static void RemoveModule(Workbook book, string moduleName)
//        {
//            VBProject vbProject = book.VBProject;
//            foreach (var comp in vbProject.VBComponents.Cast<VBComponent>().Where(comp => comp.Name == moduleName))
//            {
//                vbProject.VBComponents.Remove(comp);
//                break;
//            }
//        }
//        public static object RunModule(Workbook book, string procedureName, params object[] args)
//        {
//            return book.Application.Run(procedureName, args);
//        }
//        public static object RunModule(Workbook book, string procedureName)
//        {
//            return book.Application.Run(procedureName);
//        }
        
//    }
//}
