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

namespace CommonLibrary
{
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
        /// TODO:table index 需與sheet index 依樣
        /// (Read 多載+1) 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="isHDR">資料是否含表頭(default:true(含表頭))</param>
        /// <returns>DataSet</returns>
        public DataSet Read(string filePath,bool isHDR)
        {
            DataSet ds = new DataSet();
            using (OleDbConnection conn = new OleDbConnection(GetConnectionString(filePath, isHDR)))
            {
                conn.Open();
                try
                {
                    DataTable SchemaDT = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    foreach (DataRow r in SchemaDT.Rows)
                    {
                        using (OleDbDataAdapter dr = new OleDbDataAdapter("select * from [" + r["TABLE_NAME"].ToString() + "]", conn))
                        {
                            dr.Fill(ds, r["TABLE_NAME"].ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.Write(ex.Message);
                }
            }
            return ds;
        }

        public DataTable Read(string filePath, string sheetName)
        {
            return Read(filePath, sheetName, true);
        }
        public DataTable Read(string filePath, string sheetName, bool isHDR)
        {
            DataTable dt = new DataTable();
            dt.TableName = sheetName;
            using (OleDbConnection conn = new OleDbConnection(GetConnectionString(filePath, isHDR)))
            {
                conn.Open();
                try
                {
                    using (OleDbDataAdapter dr = new OleDbDataAdapter("select * from [" + sheetName + "$]", conn))
                    {
                        dr.Fill(dt);
                    }
                }
                catch (Exception ex)
                {
                    LogHelper.Write(ex.Message);
                }
            }
            return dt;
        }

        public void ConvertToExcel(DataTable SourceDT, string savePath)
        {
            //建立Excel
            XLWorkbook workbook = new XLWorkbook();
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
            for (int i = 1; i <= sheet.ColumnCount(); i++)
            {
                sheet.Column(i).AdjustToContents();
                if (sheet.Column(i).Width > _defaultWidth)
                    sheet.Column(i).Width = _defaultWidth;
            }
            workbook.SaveAs(savePath);
        }

        public void ConvertToExcel(DataSet SourceSet, string savePath)
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
                    for (int i = 1; i <= sheet.ColumnCount(); i++)
                    {
                        sheet.Column(i).AdjustToContents();
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

        public void ConvertToExcel2(DataSet SourceSet, string savePath)
        {
            using (XLWorkbook workbook = new XLWorkbook())
            {
                string newBomInfo = "";
                string oldBomInfo = "";
                foreach (DataRow r in SourceSet.Tables["EBOM"].Rows)
                {
                    if (r["type"].ToString() == "new")
                        newBomInfo = r["Number"].ToString() + ", " + r["Rev"].ToString() + ", " + r["RevReleaseDate"].ToString();
                    else if (r["type"].ToString() == "old")
                        oldBomInfo = r["Number"].ToString() + ", " + r["Rev"].ToString() + ", " + r["RevReleaseDate"].ToString();
                }

                foreach (DataTable SourceDT in SourceSet.Tables)
                {
                    if (SourceDT.TableName == "EBOM") {
                        continue;
                    }

                    var sheet = workbook.Worksheets.Add(SourceDT.TableName);
                    if (SourceDT.TableName == "PLM_BOM_Comparison_Report")
                    {
                        sheet.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        sheet.Style.Border.TopBorderColor = XLColor.White;
                        sheet.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                        sheet.Style.Border.RightBorderColor = XLColor.White;
                        sheet.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        sheet.Style.Border.BottomBorderColor = XLColor.White;
                        sheet.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                        sheet.Style.Border.LeftBorderColor = XLColor.White;


                        sheet.Cell(1, 2).Value = "BOM Comparison Report";
                        sheet.Cell(2, 1).Value = "Structure to Compare";
                        sheet.Cell(2, 2).Value = "BOM";
                        sheet.Cell(3, 1).Value = "Levels to Compare";
                        sheet.Cell(3, 2).Value = "All Levels";
                        sheet.Cell(4, 1).Value = "Attributes to Compare BOM by";
                        sheet.Cell(4, 2).Value = "Find Num, Item Number";
                        sheet.Cell(5, 1).Value = oldBomInfo;
                        sheet.Cell(5, 9).Value = newBomInfo;
                        sheet.Range("A5:G5").Merge();
                        sheet.Range("I5:O5").Merge();
                        var r5 = sheet.Range("A5:O5");
                        r5.Style.Fill.BackgroundColor = XLColor.FromHtml("#f0f0f0");
                        r5.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        r5.Style.Border.TopBorderColor = XLColor.LightGray;
                        r5.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                        r5.Style.Border.RightBorderColor = XLColor.LightGray;
                        r5.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        r5.Style.Border.BottomBorderColor = XLColor.LightGray;
                        r5.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                        r5.Style.Border.LeftBorderColor = XLColor.LightGray;

                        sheet.Cell("H5").Style.Fill.BackgroundColor = XLColor.LightGray;

                        foreach (DataColumn col in SourceDT.Columns)
                        {
                            var header = sheet.Cell(6, col.Ordinal + 1);
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

                            if (col.ColumnName == "OldLevel")
                                header.Value = "Level";
                            else if (col.ColumnName == "OldNumber")
                                header.Value = "Number";
                            else if (col.ColumnName == "OldQty")
                                header.Value = "Qty";
                            else if (col.ColumnName == "FindNum" || col.ColumnName == "OldFindNum")
                                header.Value = "Find Num";
                            else if (col.ColumnName == "ItemDescription" || col.ColumnName == "OldItemDescription")
                                header.Value = "Item Description";
                            else if (col.ColumnName == "RefDes" || col.ColumnName == "OldRefDes")
                                header.Value = "Ref Des";
                            else if (col.ColumnName == "BomNotes" || col.ColumnName == "OldBomNotes")
                                header.Value = "BOM Notes";
                            else if (col.ColumnName == "space")
                            {
                                header.Value = "";
                                header.Style.Fill.BackgroundColor = XLColor.LightGray;
                            }
                            else
                                header.Value = col.ColumnName;
                        }

                        int rowIdx = 7;
                        foreach (DataRow item in SourceDT.Rows)
                        {
                            foreach (DataColumn col in SourceDT.Columns)
                            {
                                var cell = sheet.Cell(rowIdx, col.Ordinal + 1)
;                               cell.Value = item[col.Ordinal];
                                cell.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                                cell.Style.Border.TopBorderColor = XLColor.LightGray;
                                cell.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                                cell.Style.Border.RightBorderColor = XLColor.LightGray;
                                cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                                cell.Style.Border.BottomBorderColor = XLColor.LightGray;
                                cell.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                                cell.Style.Border.LeftBorderColor = XLColor.LightGray;
                            }
                            sheet.Row(rowIdx).Height = 30;
                            sheet.Cell(rowIdx, 8).Style.Fill.BackgroundColor = XLColor.LightGray;
                            if (sheet.Cell(rowIdx, 1).Value.ToString() == "")
                                sheet.Range("A:G").Row(rowIdx).Style.Fill.BackgroundColor = XLColor.FromHtml("#c0c0c0");
                            else if (sheet.Cell(rowIdx, 9).Value.ToString() == "")
                                sheet.Range("I:O").Row(rowIdx).Style.Fill.BackgroundColor = XLColor.FromHtml("#c0c0c0");
                            else
                            {
                                sheet.Cell(rowIdx, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#c0c0c0");
                                if (sheet.Cell(rowIdx, 6).Value.ToString() != "")
                                    sheet.Cell(rowIdx, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#c0c0c0");
                                sheet.Cell(rowIdx, 13).Style.Fill.BackgroundColor = XLColor.FromHtml("#c0c0c0");
                                if (sheet.Cell(rowIdx, 14).Value.ToString() != "")
                                    sheet.Cell(rowIdx, 14).Style.Fill.BackgroundColor = XLColor.FromHtml("#c0c0c0");
                            }

                            rowIdx++;
                        }

                    }
                    else if (SourceDT.TableName == "PartNumber_Change")
                    {
                        sheet.Cell(1,1).Value = oldBomInfo;
                        sheet.Cell(1,3).Value = newBomInfo;
                        sheet.Range("A:B").Row(1).Merge();
                        sheet.Range("C:D").Row(1).Merge();
                        foreach (DataColumn col in SourceDT.Columns)
                        {
                            //var header = sheet.Range("A1:C1");
                            var header = sheet.Cell(2, col.Ordinal + 1);
                            if (col.ColumnName == "OldNumber")
                                header.Value = "Number";
                            else if (col.ColumnName == "OldItemDescription")
                                header.Value = "Item Description";
                            else if (col.ColumnName == "ItemDescription")
                                header.Value = "Item Description";
                            else if (col.ColumnName == "RefDes")
                                header.Value = "Ref Des";
                            else if (col.ColumnName == "BomNotes")
                                header.Value = "BOM Notes";
                            else
                                header.Value = col.ColumnName;
                            header.Style.Fill.BackgroundColor = XLColor.FromHtml("#f0f0f0");
                            header.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                            header.Style.Border.TopBorderColor = XLColor.LightGray;
                            header.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                            header.Style.Border.RightBorderColor = XLColor.LightGray;
                            header.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                            header.Style.Border.BottomBorderColor = XLColor.LightGray;
                            header.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                            header.Style.Border.LeftBorderColor = XLColor.LightGray;
                        }

                        int rowIdx = 3;
                        foreach (DataRow item in SourceDT.Rows)
                        {
                            foreach (DataColumn col in SourceDT.Columns)
                            {
                                sheet.Cell(rowIdx, col.Ordinal + 1).Value = item[col.Ordinal];
                            }
                            rowIdx++;
                        }
                    }
                    else
                    {
                        foreach (DataColumn col in SourceDT.Columns)
                        {
                            //var header = sheet.Range("A1:C1");
                            var header = sheet.Cell(1, col.Ordinal + 1);
                            if (col.ColumnName == "FindNum")
                                header.Value = "Find Num";
                            else if (col.ColumnName == "ItemDescription")
                                header.Value = "Item Description";
                            else if (col.ColumnName == "RefDes")
                                header.Value = "Ref Des";
                            else if (col.ColumnName == "BomNotes")
                                header.Value = "BOM Notes";
                            else
                                header.Value = col.ColumnName;
                            header.Style.Fill.BackgroundColor = XLColor.FromHtml("#f0f0f0");
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
                    }

                    //文字置頂
                    sheet.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    //自動換行
                    sheet.Style.Alignment.WrapText = true;
                    //自動伸縮欄寬
                    for (int i = 1; i <= sheet.ColumnCount(); i++)
                    {
                        sheet.Column(i).AdjustToContents();
                        if (sheet.Column(i).Width > _defaultWidth)
                            sheet.Column(i).Width = _defaultWidth;
                    }
                }
                workbook.SaveAs(savePath);
            }
        }

        //public void ConvertToExcel(XLWorkbook workbook, string savePath)
        //{
        //    try
        //    {
        //        workbook.SaveAs(savePath);
        //    }
        //    catch (Exception)
        //    {
        //    }
        //    finally
        //    {
        //        workbook.Dispose();
        //    }
        //}

        //public XLWorkbook ConvertToWorkbook() {
        //}

        public DataTable ConvertToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
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
