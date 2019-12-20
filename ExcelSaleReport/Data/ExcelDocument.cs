using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;

namespace ExcelSaleReport.Data
{
    public class ExcelDocument
    {
        private string fontName = "Arial";
        private int fontSize = 9;
        private static string excelDir = Path.GetFullPath(Path.Combine(Application.StartupPath + "\\Reports\\"));
        private Microsoft.Office.Interop.Excel._Worksheet _workSheet;
        private Microsoft.Office.Interop.Excel.Application _excelApp;

        public ExcelDocument()
        {
            _excelApp = new Microsoft.Office.Interop.Excel.Application();
            if(_excelApp!=null)
                _excelApp.Workbooks.Add();
        }

        public void CreateSheet(string sheetName, int sheetNumber, string title, string header, DataTable dtExcel)
        {
            try
            {
                if (_excelApp == null)
                {
                    MessageBox.Show("Excel not insalled...", "Attention!");
                    return;
                }

                _workSheet = _excelApp.Worksheets.Add();
                _workSheet.Name = sheetName;
                _workSheet.Rows.Font.Name = fontName;
                _workSheet.Rows.Font.Size = fontSize;
                
                _workSheet.Cells[1, 1] = title;
                _workSheet.Cells[3, 1] = header;
                
                _workSheet.Range[_workSheet.Cells[1, 1], _workSheet.Cells[1, dtExcel.Columns.Count]].Merge();
                _workSheet.Range[_workSheet.Cells[2, 1], _workSheet.Cells[2, dtExcel.Columns.Count]].Merge();
                _workSheet.Range[_workSheet.Cells[3, 1], _workSheet.Cells[3, dtExcel.Columns.Count]].Merge();
                _workSheet.Range[_workSheet.Cells[4, 1], _workSheet.Cells[4, dtExcel.Columns.Count]].Merge();
                
                _workSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                _workSheet.Cells[3, 1].EntireRow.Font.Bold = true;

                InsertColumns(_workSheet, dtExcel);
                InsertRows(_workSheet, dtExcel, 6);

                _workSheet.Tab.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5;
               /* _workSheet = (Microsoft.Office.Interop.Excel.Worksheet)_excelApp.Worksheets.get_Item(1);
                _workSheet.Select();*/
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error - Excel! \n Exception: "+ex.Message);
            }
        }

        public void CreateSheetWithMultiLayerColumn(string sheetName, int sheetNumber, string title, string header, DataTable dtExcel)
        {
            try
            {
                _workSheet = (Microsoft.Office.Interop.Excel.Worksheet)_excelApp.Worksheets.get_Item(sheetNumber);
                _workSheet.Select();
                _workSheet.Name = sheetName;
                
                _workSheet.Rows.Font.Name = fontName;
                _workSheet.Rows.Font.Size = fontSize;
                
                _workSheet.Cells[1, 1] = title;
                _workSheet.Cells[3, 1] = header;
                _workSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                _workSheet.Cells[3, 1].EntireRow.Font.Bold = true;
                _workSheet.Range[_workSheet.Cells[1, 1], _workSheet.Cells[1, dtExcel.Columns.Count]].Merge();
                _workSheet.Range[_workSheet.Cells[2, 1], _workSheet.Cells[2, dtExcel.Columns.Count]].Merge();
                _workSheet.Range[_workSheet.Cells[3, 1], _workSheet.Cells[3, dtExcel.Columns.Count]].Merge();
                _workSheet.Range[_workSheet.Cells[4, 1], _workSheet.Cells[4, dtExcel.Columns.Count]].Merge();

                InsertColumnsMultiLayer(_workSheet, dtExcel);
                InsertRows(_workSheet, dtExcel, 7);
                _workSheet.Tab.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorAccent5;
                _workSheet = (Microsoft.Office.Interop.Excel.Worksheet)_excelApp.Worksheets.get_Item(1);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void SaveExcelDoc(string fileName)
        {
            string path = excelDir + fileName + "-" + (new Random().Next(1, 100)).ToString() + ".XLSX";
            _workSheet.SaveAs(path);
            _excelApp.Quit();
            System.Diagnostics.Process.Start(path);
        }

        private void InsertRows(Microsoft.Office.Interop.Excel._Worksheet workSheet, DataTable dtExcel, int startRow)
        {
            int k = startRow;
            int t = 1;
            for (var i = 0; i < dtExcel.Rows.Count; i++)
            {
                for (var j = 0; j < dtExcel.Columns.Count; j++)
                {
                    workSheet.Cells[k, t] = dtExcel.Rows[i][j];
                    t = t + 1;
                }
                k = k + 1;
                t = 1;
                workSheet.Columns.AutoFit();
            }
        }

        private void InsertColumns(Microsoft.Office.Interop.Excel._Worksheet workSheet, DataTable dtExcel)
        {
            int k = 5;
            int t = 2;
            workSheet.Cells[k, 1] = dtExcel.Columns[0].ColumnName.ToUpper();
            workSheet.Cells[k, 2] = dtExcel.Columns[1].ColumnName.ToUpper();

            for (int i = 2; i <= dtExcel.Columns.Count - 2; i++)
            {
                t = t + 1;
                workSheet.Cells[k, t] = "-" + dtExcel.Columns[i].ColumnName.ToString().ToUpper() + "-";
                workSheet.Columns[t].NumberFormat = "0,00";
            }

            workSheet.Cells[k, t + 1] = dtExcel.Columns[dtExcel.Columns.Count - 1].ColumnName.ToString().ToUpper();
            workSheet.Columns[t + 1].NumberFormat = "0,00";

            var columnHeading = workSheet.Range[
                workSheet.Cells[5, 1],
                workSheet.Cells[5, dtExcel.Columns.Count]];
            columnHeading.Font.Bold = true;
            columnHeading.Font.Color = Color.White;
            columnHeading.Interior.Color = Color.RoyalBlue;
            columnHeading.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            var columnBorder = workSheet.Range[
                workSheet.Cells[5, 1],
                workSheet.Cells[5 + dtExcel.Rows.Count, dtExcel.Columns.Count]];
            columnBorder.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            columnBorder.Borders.Color = Color.Black;

            var columnHeadingAligment = workSheet.Range[
                workSheet.Cells[5, 3],
                workSheet.Cells[5, dtExcel.Columns.Count]];
            columnHeadingAligment.Cells.HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
        }
        private void InsertColumnsMultiLayer(Microsoft.Office.Interop.Excel._Worksheet workSheet, DataTable dtExcel)
        {
            int k = 5;
            int t = 3;
            int days = 31;

            workSheet.Cells[k, 1] = "Id";
            workSheet.Cells[k, 2] = "Name";

            workSheet.Range[workSheet.Cells[k, 1], workSheet.Cells[k + 1, 1]].Merge();
            workSheet.Range[workSheet.Cells[k, 2], workSheet.Cells[k + 1, 2]].Merge();

            workSheet.Range[workSheet.Cells[k, 1], workSheet.Cells[k + 1, 1]].Cells.VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            workSheet.Range[workSheet.Cells[k, 2], workSheet.Cells[k + 1, 2]].Cells.VerticalAlignment =
                Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            for (int i = 1; i <= days; i++)
            {
                workSheet.Cells[k, i * 3] = i;
                workSheet.Range[workSheet.Cells[k, i * 3], workSheet.Cells[k, i * 3 + 2]].Merge();
                workSheet.Range[workSheet.Cells[k, i * 3], workSheet.Cells[k, i * 3 + 2]].Cells.HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                workSheet.Cells[k + 1, t] = "% DIP";//Differnce in price %
                workSheet.Cells[k + 1, t].Cells.HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[k + 1, t + 1] = "DIP Amount"; // Differnce in price amount
                workSheet.Cells[k + 1, t + 1].Cells.HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Cells[k + 1, t + 2] = "RV"; //Retail Value
                workSheet.Cells[k + 1, t + 2].Cells.HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                workSheet.Columns[t].NumberFormat = "0,00";
                workSheet.Columns[t + 1].NumberFormat = "0,00";
                workSheet.Columns[t + 2].NumberFormat = "0,00";

                t = t + 3;
            }

            workSheet.Cells[k, (days + 1) * 3] = "Total";
            workSheet.Cells[k + 1, (days + 1) * 3] = "% DIP";
            workSheet.Cells[k + 1, (days + 1) * 3 + 1] = "DIP Amount";
            workSheet.Cells[k + 1, (days + 1) * 3 + 2] = "RV";
            workSheet.Columns[(days + 1) * 3].NumberFormat = "0,00";
            workSheet.Columns[(days + 1) * 3 + 1].NumberFormat = "0,00";
            workSheet.Columns[(days + 1) * 3 + 2].NumberFormat = "0,00";

            workSheet.Cells[k + 1, (days + 1) * 3].Cells.HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            workSheet.Cells[k + 1, (days + 1) * 3 + 1].Cells.HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            workSheet.Cells[k + 1, (days + 1) * 3 + 2].Cells.HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            workSheet.Range[workSheet.Cells[k, (days + 1) * 3], workSheet.Cells[k, (days + 1) * 3 + 2]].Merge();
            workSheet.Range[workSheet.Cells[k, (days + 1) * 3], workSheet.Cells[k, (days + 1) * 3 + 2]].Cells.HorizontalAlignment =
                 Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

            var columnHeading = workSheet.Range[
                workSheet.Cells[k, 1],
                workSheet.Cells[k + 1, dtExcel.Columns.Count]];
            columnHeading.Font.Bold = true;
            columnHeading.Font.Color = Color.White;
            columnHeading.Interior.Color = Color.RoyalBlue;
            var columnBorder = workSheet.Range[
                workSheet.Cells[5, 1],
                workSheet.Cells[6 + dtExcel.Rows.Count, dtExcel.Columns.Count]];
            columnBorder.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            columnBorder.Borders.Color = Color.Black;
        }
    }
}
