using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace ExcelToolkitAddIn
{
    #region PANE
    /// <summary>
    /// PANE 基础页面交互后端类
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class PANE
    {
        public Layout layout = Globals.ThisAddIn.layout;
        public Sqlite sql;
        public PANE()
        {
            sql = layout.sql;
        }
    }
    #endregion

    #region Query
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Query : PANE
    {
        public string[] GetLevel1()
        {
            const string sqlCode = @"SELECT MIN(id) as id, level1 FROM 内容 GROUP BY level1 ORDER BY id ASC";
            var list = sql.GetData(sqlCode);
            string[] level1 = new string[list.Count];
            for (int i = 0; i < list.Count; i++) level1[i] = list[i]["level1"].ToString();
            return level1;
        }

        public string GetLevel2(string level1)
        {
            string sqlCode = $"SELECT level1,level2,info FROM 内容 WHERE level1 = '{level1}'";
            return sql.DataToJson(sql.GetData(sqlCode));
        }

        public string Search(string key, bool content = true, bool info = true, bool level2 = true)
        {
            if (string.IsNullOrEmpty(key) || !(content || info || level2))
                return "";
            string sqlCode = "SELECT level1,level2,info FROM 内容 WHERE 1>1";
            if (content) sqlCode += $" OR content LIKE '%{key}%'";
            if (info) sqlCode += $" OR info LIKE '%{key}%'";
            if (level2) sqlCode += $" OR level2 LIKE '%{key}%'";
            return sql.DataToJson(sql.GetData(sqlCode));
        }

        public string GetContent(string level1, string level2)
        {
            var sqlCode = $"SELECT content FROM 内容 WHERE level1 = '{level1}' AND level2 = '{level2}'";
            var list = sql.GetData(sqlCode);
            return list[0]["content"].ToString();
        }
    }
    #endregion

    #region Name
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Name : PANE
    {
        public string GetTable()
        {
            var data = sql.GetTableAll("名称管理器");
            foreach (var row in data)
            {
                row["isInApp"] = row["isInApp"].ToString() == "1";
                row["isInBook"] = row["isInBook"].ToString() == "1";
                row["isInSheet"] = row["isInSheet"].ToString() == "1";
            }

            return sql.DataToJson(data);
        }

        public void Insert(string json)
        {

        }

        public int Remove(params string[] ids)
        {
            return sql.Remove("名称管理器", ids);
        }

        public int Update(string id, string json)
        {
            Dictionary<string, string> data =
                JsonSerializer
                .Deserialize<Dictionary<string, object>>(json)
                .ToDictionary(kvp => kvp.Key, kvp => kvp.Value.ToString());
            return sql.UpDate("名称管理器", data, $"id = '{id}'");
        }
    }
    #endregion

    #region PowerQuery
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class PowerQuery : PANE
    {
        private string[] columns = { "name", "code" };
        private string tableName = "PowerQuery";
        #region Sqlite交互部分
        public string GetTable()
        {
            return sql.DataToJson(sql.GetTableAll(tableName));
        }

        public void Insert(string name, string code)
        {
            string[] data = { name, code };
            sql.Insert(tableName, this.columns, data);
        }

        public int Remove(string id)
        {
            return sql.Remove(tableName, $"id='{id}'");
        }

        public void UpDate(string id, string name, string code)
        {
            sql.UpDate(tableName, $"name='{name}',code='{code}'", $"id='{id}'");
        }

        public int GetMaxId()
        {
            return sql.GetMaxId(tableName);
        }
        #endregion

        #region Excel交互部分
        public void InsertPq(string name, string code)
        {
            string safeName = name.Replace("\"", "\"\"");
            string safeCode = code.Replace("\"", "\"\"");

            dynamic book = layout.app.ActiveWorkbook;
            dynamic queries = book.Queries;

            dynamic existing = null;
            foreach (dynamic q in queries)
            {
                if (q.Name != safeName) continue;
                existing = q;
                break;
            }

            if (existing != null)
            {
                existing.Formula = safeCode;
            }
            else
            {
                queries.Add(Name: safeName, Formula: safeCode);
            }
        }

        public string ReadPQ()
        {
            dynamic book = layout.app.ActiveWorkbook;
            dynamic query = book.Queries;
            int count = query.Count;
            if (count == 0) return "[]";
            string[][] output = new string[count][];
            int i = 0;
            foreach (dynamic q in query)
            {
                output[i] = new string[] { q.Name, q.Formula };
                i++;
            }

            return JsonSerializer.Serialize(output);
        }
        public string ReadAllPQ()
        {
            dynamic app = layout.app;
            List<string[]> allQueries = new List<string[]>();
            foreach (dynamic workbook in app.Workbooks)
            {
                try
                {
                    dynamic queries = workbook.Queries;
                    int count = queries.Count;
                    if (count <= 0) continue;
                    foreach (dynamic query in queries)
                    {
                        allQueries.Add(new string[]
                        {
                            query.Name,
                            query.Formula
                        });
                    }
                }
                catch { continue; }
            }

            if (allQueries.Count > 0)
            {
                return JsonSerializer.Serialize(allQueries);
            }
            return "[]";
        }

        #region 调用PowerQuery备用Vba方法
        //public void InsertPqVba(string name, string code)
        //{
        //    name = name.Replace("\"", "\"\"");
        //    code = code.Replace("\"", "\"\"");
        //    Workbook book = layout.app.ActiveWorkbook;
        //    string vbaCode = $@"
        //                        Public Sub UpsertPQ()
        //                            found = False
        //                            For Each qry In ThisWorkbook.Queries
        //                                If qry.Name = ""{name}"" Then
        //                                    qry.Formula = ""{code}""
        //                                    found = True
        //                                    Exit For
        //                                End If
        //                            Next
        //                            If Not found Then
        //                                ThisWorkbook.Queries.Add Name:=""{name}"", Formula:=""{code}""
        //                            End If
        //                        End Sub";
        //    Vba.AddModule(book, "tempPQ", vbaCode);
        //    Vba.RunModule(book, "UpsertPQ");
        //    Vba.RemoveModule(book, "tempPQ");
        //}
        //public string ReadPQVba()
        //{
        //    Workbook book = layout.app.ActiveWorkbook;
        //    const string vbaCode = @"
        //                        Public Function GetPowerQueryMCode() As Variant
        //                            On Error GoTo ErrorHandler
        //                            Dim queriesCount As Long
        //                            queriesCount = ThisWorkbook.Queries.Count
        //                            If queriesCount = 0 Then
        //                                GetPowerQueryMCode = Array()
        //                                Exit Function
        //                            End If
        //                            Dim result() As Variant
        //                            ReDim result(1 To queriesCount, 1 To 2)
        //                            Dim i As Long
        //                            Dim qry As WorkbookQuery
        //                            For i = 1 To queriesCount
        //                                Set qry = ThisWorkbook.Queries(i)
        //                                result(i, 1) = qry.Name
        //                                result(i, 2) = qry.Formula
        //                            Next i
        //                            GetPowerQueryMCode = result
        //                            Exit Function
        //                        ErrorHandler:
        //                            GetPowerQueryMCode = Array()
        //                        End Function";
        //    Vba.AddModule(book, "temp", vbaCode);
        //    object result = Vba.RunModule(book, "GetPowerQueryMCode");
        //    Vba.RemoveModule(book, "temp");
        //    if (!(result is object[,] arr)) return "[]";
        //    int rows = arr.GetLength(0);
        //    int cols = arr.GetLength(1);

        //    string[][] output = new string[rows][];
        //    for (int i = 0; i < rows; i++)
        //    {
        //        output[i] = new string[cols];
        //        for (int j = 0; j < cols; j++)
        //        {
        //            output[i][j] = arr[i + 1, j + 1]?.ToString() ?? "";
        //        }
        //    }
        //    return JsonSerializer.Serialize(output);
        //}


        #endregion

        #endregion
    }
    #endregion


}
