using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using CommonLibrary.Extensions;
using ClosedXML.Excel;
using System.ComponentModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace CommonLibrary
{
    /// <summary>
    /// excel存取元件
    /// </summary>
    public class ExcelHelper
    {
        private int _defaultWidth = 50;
      
        /// <summary>
        /// 讀取excel檔並輸出成 Dataset
        /// </summary>
        /// <param name="filePath">檔案路徑</param>
        /// <returns>DataSet</returns>
        public DataSet Read(string filePath)
        {
            return Read(filePath, true);
        }
        /// <summary>
        /// (Read 多載+1) 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="isHDR">資料是否含表頭(default:true(含表頭))</param>
        /// <returns>DataSet</returns>
        public DataSet Read(string filePath, bool isHDR)
        {
            DataSet ds = new DataSet();
            #region test code
            //#OleDB
            //using (OleDbConnection conn = new OleDbConnection(GetConnectionString(filePath, isHDR)))
            //{
            //    conn.Open();
            //    try
            //    {
            //        DataTable SchemaDT = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //        foreach (DataRow r in SchemaDT.Rows)
            //        {
            //            using (OleDbDataAdapter dr = new OleDbDataAdapter("select * from [" + r["TABLE_NAME"].ToString() + "]", conn))
            //            {
            //                dr.Fill(ds, r["TABLE_NAME"].ToString());
            //            }
            //        }
            //    }
            //    catch (Exception ex)
            //    {
            //        LogHelper.Write(ex.Message);
            //    }
            //}

            //#linqtoexcel
            //ExcelQueryFactory excel = new ExcelQueryFactory(filePath);
            //DataTable dt;
            //foreach (var sheetname in excel.GetWorksheetNames())
            //{
            //    dt = new DataTable();
            //    if (isHDR)
            //    {
            //        var worksheet = excel.Worksheet(sheetname);
            //        var query = from row in worksheet select row;
            //        var columnName = excel.GetColumnNames(sheetname);
            //        //建立欄位名稱
            //        foreach (var col in columnName)
            //            dt.Columns.Add(col.ToString());
            //        //寫入資料到資料列
            //        foreach (Row item in query)
            //        {
            //            dt.NewRow();
            //            object[] cell = new object[columnName.Count()];
            //            int idx = 0;
            //            foreach (var col in columnName)
            //            {
            //                cell[idx] = item[col].Value;
            //                idx++;
            //            }
            //            dt.Rows.Add(cell);
            //        }
            //    }
            //    else {
            //        var worksheet = excel.WorksheetNoHeader(sheetname);
            //        var query = from row in worksheet select row;
            //        var columnName = excel.GetColumnNames(sheetname);
            //        //建立欄位名稱
            //        foreach (var col in columnName)
            //            dt.Columns.Add("");
            //        //寫入資料到資料列
            //        foreach (RowNoHeader item in query)
            //        {
            //            dt.NewRow();
            //            object[] cell = new object[columnName.Count()];
            //            int idx = 0;
            //            for (int i = 0; i < columnName.Count(); i++)
            //            {
            //                cell[idx] = item[i].Value;
            //                idx++;
            //            }
            //            dt.Rows.Add(cell);
            //        }
            //    }
            //    ds.Tables.Add(dt);
            //}
            #endregion

            IWorkbook wb;
            try
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    FileInfo info = new FileInfo(filePath);
                    if (info.Extension == ".xlsx")
                        wb = new XSSFWorkbook(file);
                    else
                        wb = new HSSFWorkbook(file);
                }

                DataTable dt;
                for (int i = 0; i < wb.NumberOfSheets; ++i)
                {
                    dt = new DataTable();
                    ISheet sheet = wb.GetSheetAt(i);
                    dt.TableName = sheet.SheetName;
                    int firstRowindex = 0;
                    int colMaxCount = sheet.GetRow(firstRowindex) == null ? 1 : sheet.GetRow(firstRowindex).LastCellNum;
                    if (isHDR)
                    {
                        for (int c = 0; c < colMaxCount; c++)
                            dt.Columns.Add((sheet.GetRow(firstRowindex) == null || sheet.GetRow(firstRowindex).GetCell(c) == null)
                                ? "" : sheet.GetRow(firstRowindex).GetCell(c).StringCellValue);
                        firstRowindex++;
                    }
                    else
                    {
                        for (int c = 0; c < colMaxCount; c++)
                            dt.Columns.Add("");
                    }

                    for (int r = firstRowindex; r <= sheet.LastRowNum; r++)
                    {
                        //columns ++
                        if (sheet.GetRow(r) != null && sheet.GetRow(r).LastCellNum > colMaxCount)
                        {
                            colMaxCount = sheet.GetRow(r).LastCellNum;
                            for (int c = dt.Columns.Count; c < colMaxCount; c++)
                                dt.Columns.Add("");
                        }

                        dt.NewRow();
                        object[] cell = new object[colMaxCount];
                        int idx = 0;
                        for (int c = 0; c < colMaxCount; c++)
                        {
                            cell[idx] = (sheet.GetRow(r) != null) ? sheet.GetRow(r).GetCell(c) : null;
                            idx++;
                        }
                        dt.Rows.Add(cell);

                    }
                    ds.Tables.Add(dt);
                }
            }
            catch (Exception ex)
            {
                LogHelper.Write("[Error][CommonLibrary>> ExcelHelper>> Read]\n" + ex.Message);
                throw new Exception("[CommonLibrary]ExcelHelper Read failed!!");
            }
            return ds;
        }

        /// <summary>
        /// 讀取excel檔單一sheet
        /// </summary>
        /// <param name="filePath">檔案路徑</param>
        /// <param name="sheetName">sheet名稱</param>
        /// <returns>DataTable</returns>
        public DataTable Read(string filePath, string sheetName)
        {
            return Read(filePath, sheetName, true);
        }
        /// <summary>
        /// (Read 多載+1) 
        /// </summary>
        /// <param name="filePath">檔案路徑</param>
        /// <param name="sheetName">sheet名稱</param>
        /// <param name="isHDR"></param>
        /// <returns>DataTable</returns>
        public DataTable Read(string filePath, string sheetName, bool isHDR)
        {
            DataTable dt = new DataTable();

            IWorkbook wb;
            try
            {
                using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    FileInfo info = new FileInfo(filePath);
                    if (info.Extension == ".xlsx")
                        wb = new XSSFWorkbook(file);
                    else
                        wb = new HSSFWorkbook(file);
                }

                ISheet sheet = wb.GetSheet(sheetName);
                dt.TableName = sheet.SheetName;
                int firstRowindex = 0;
                int colMaxCount = sheet.GetRow(firstRowindex) == null ? 1 : sheet.GetRow(firstRowindex).LastCellNum;
                if (isHDR)
                {
                    for (int c = 0; c < colMaxCount; c++)
                        dt.Columns.Add((sheet.GetRow(firstRowindex) == null || sheet.GetRow(firstRowindex).GetCell(c) == null)
                            ? "" : sheet.GetRow(firstRowindex).GetCell(c).StringCellValue);
                    firstRowindex++;
                }
                else
                {
                    for (int c = 0; c < colMaxCount; c++)
                        dt.Columns.Add("");
                }

                for (int r = firstRowindex; r <= sheet.LastRowNum; r++)
                {
                    //columns ++
                    if (sheet.GetRow(r) != null && sheet.GetRow(r).LastCellNum > colMaxCount)
                    {
                        colMaxCount = sheet.GetRow(r).LastCellNum;
                        for (int c = dt.Columns.Count; c < colMaxCount; c++)
                            dt.Columns.Add("");
                    }

                    dt.NewRow();
                    object[] cell = new object[colMaxCount];
                    int idx = 0;
                    for (int c = 0; c < colMaxCount; c++)
                    {
                        cell[idx] = (sheet.GetRow(r) != null) ? sheet.GetRow(r).GetCell(c) : null;
                        idx++;
                    }
                    dt.Rows.Add(cell);
                }
            }
            catch (Exception ex)
            {
                LogHelper.Write("[Error][CommonLibrary>> ExcelHelper>> Read]\n" + ex.Message);
                throw new Exception("[CommonLibrary]ExcelHelper Read failed!!");
            }
            return dt;
        }

        /// <summary>
        /// DataTable 轉出 excel
        /// </summary>
        /// <param name="SourceDT">DataTable</param>
        /// <param name="savePath">輸出路徑</param>
        public void ConvertToExcel(DataTable SourceDT, string savePath)
        {
            try
            {
                //建立Excel
                using (XLWorkbook workbook = new XLWorkbook())
                {
                    //XLWorkbook workbook = new XLWorkbook();
                    var sheet = workbook.Worksheets.Add(string.IsNullOrEmpty(SourceDT.TableName) ? "Sheet1" : SourceDT.TableName);
                    foreach (DataColumn col in SourceDT.Columns)
                    {
                        //var header = sheet.Range("A1:C1");
                        var header = sheet.Cell(1, col.Ordinal + 1);
                        header.Value = col.ColumnName;
                        header.Style.Fill.BackgroundColor = XLColor.FromHtml("#f0f0f0");
                        //header.Style.Font.FontColor = XLColor.Yellow;
                        //header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        header.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        header.Style.Border.TopBorderColor = XLColor.LightGray;
                        header.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                        header.Style.Border.RightBorderColor = XLColor.LightGray;
                        header.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        header.Style.Border.BottomBorderColor = XLColor.LightGray;
                        header.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                        header.Style.Border.LeftBorderColor = XLColor.LightGray;
                    }

                    int rowIdx = 2;
                    foreach (DataRow item in SourceDT.Rows)
                    {
                        foreach (DataColumn col in SourceDT.Columns)
                        {
                            sheet.Cell(rowIdx, col.Ordinal + 1).Value = item[col.Ordinal];
                        }
                        rowIdx++;
                    }

                    //文字置頂
                    sheet.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    //自動換行
                    sheet.Style.Alignment.WrapText = true;
                    //自動伸縮欄寬
                    sheet.ColumnsUsed().AdjustToContents();
                    for (int i = 1; i <= sheet.ColumnsUsed().Count(); i++)
                    {
                        if (sheet.Column(i).Width > _defaultWidth)
                            sheet.Column(i).Width = _defaultWidth;
                    }

                    workbook.SaveAs(savePath);
                }
            }
            catch (Exception ex)
            {
                LogHelper.Write("[Error][CommonLibrary>> ExcelHelper>> ConvertToExcel]\n" + ex.Message);
                throw new Exception("[CommonLibrary]ExcelHelper ConvertToExcel failed!!");
            }
        }
        /// <summary>
        /// DataSet 轉出 excel
        /// </summary>
        /// <param name="SourceSet">DataSet</param>
        /// <param name="savePath">輸出路徑</param>
        public void ConvertToExcel(DataSet SourceSet, string savePath)
        {
            try
            {
                using (XLWorkbook workbook = new XLWorkbook())
                {
                    foreach (DataTable SourceDT in SourceSet.Tables)
                    {
                        var sheet = workbook.Worksheets.Add(SourceDT.TableName);

                        foreach (DataColumn col in SourceDT.Columns)
                        {
                            //var header = sheet.Range("A1:C1");
                            var header = sheet.Cell(1, col.Ordinal + 1);
                            header.Value = col.ColumnName;
                            header.Style.Fill.BackgroundColor = XLColor.FromHtml("#f0f0f0");
                            //header.Style.Font.FontColor = XLColor.Yellow;
                            //header.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            header.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            header.Style.Border.TopBorderColor = XLColor.LightGray;
                            header.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                            header.Style.Border.RightBorderColor = XLColor.LightGray;
                            header.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            header.Style.Border.BottomBorderColor = XLColor.LightGray;
                            header.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                            header.Style.Border.LeftBorderColor = XLColor.LightGray;
                        }

                        int rowIdx = 2;
                        foreach (DataRow item in SourceDT.Rows)
                        {
                            foreach (DataColumn col in SourceDT.Columns)
                            {
                                sheet.Cell(rowIdx, col.Ordinal + 1).Value = item[col.Ordinal];
                            }
                            rowIdx++;
                        }

                        //文字置頂
                        sheet.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                        //自動換行
                        sheet.Style.Alignment.WrapText = true;
                        //自動伸縮欄寬
                        sheet.ColumnsUsed().AdjustToContents();
                        for (int i = 1; i <= sheet.ColumnsUsed().Count(); i++)
                        {
                            if (sheet.Column(i).Width > _defaultWidth)
                                sheet.Column(i).Width = _defaultWidth;
                        }

                        //TODO:Heigt取不到值
                        //if (sheet.Row(10).Height > 30)
                        //    sheet.Row(10).Height = 30;

                    }
                    workbook.SaveAs(savePath);
                }
            }
            catch (Exception ex)
            {
                LogHelper.Write("[Error][CommonLibrary>> ExcelHelper>> ConvertToExcel]\n" + ex.Message);
                throw new Exception("[CommonLibrary]ExcelHelper ConvertToExcel failed!!");
            }
        }
        /// <summary>
        /// List 轉 DataTable
        /// TODO:可移出excel class使用
        /// </summary>
        /// <typeparam name="T">data struct</typeparam>
        /// <param name="data">data</param>
        /// <returns>DataTable</returns>
        public DataTable ConvertToDataTable<T>(IList<T> data)
        {
            DataTable table = new DataTable();
            try
            {
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
                foreach (PropertyDescriptor prop in properties)
                    table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                foreach (T item in data)
                {
                    DataRow row = table.NewRow();
                    foreach (PropertyDescriptor prop in properties)
                        row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                    table.Rows.Add(row);
                }
            }
            catch (Exception ex)
            {
                LogHelper.Write("[Error][CommonLibrary>> ExcelHelper>> ConvertToDataTable]\n" + ex.Message);
                throw new Exception("[CommonLibrary]ExcelHelper ConvertToDataTable failed!!");
            }
            return table;
        }

        //public MemoryStream ConvertToMemory(DataSet SourceSet)
        //{
        //    MemoryStream ms = new MemoryStream();
        //    XLWorkbook workbook = new XLWorkbook();
        //    foreach (DataTable SourceDT in SourceSet.Tables)
        //    {
        //        var sheet = workbook.Worksheets.Add(SourceDT.TableName);
        //        foreach (DataColumn col in SourceDT.Columns)
        //            sheet.Cell(1, col.Ordinal + 1).Value = col.ColumnName;
        //        int rowIdx = 2;
        //        foreach (DataRow item in SourceDT.Rows)
        //        {
        //            foreach (DataColumn col in SourceDT.Columns)
        //            {
        //                sheet.Cell(rowIdx, col.Ordinal + 1).Value = item[col.Ordinal];
        //            }
        //            rowIdx++;
        //        }
        //    }
        //    workbook.SaveAs(ms);
        //    return ms;
        //}

        private string GetConnectionString(string filePath, bool isHDR)
        {
            string HDR = isHDR ? "YES" : "NO";//資料是否含表頭
            if (string.Compare(Path.GetExtension(filePath), ".xlsx", true) == 0)
                return @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + filePath + ";Extended Properties='Excel 12.0;HDR="+ HDR + "'";
            return @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + filePath + ";Extended Properties='Excel 8.0;HDR=" + HDR + ";IMEX=1'";
        }
    }

  
}
