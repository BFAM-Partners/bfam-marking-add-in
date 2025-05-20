using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarkingSheet
{
    internal class Utils
    {

        public class DataTableContent
        {
            public string Name { get; set; }
            public List<string> Headers = new List<string>();
            public List<Dictionary<string, object>> Rows = new List<Dictionary<string, object>>();
        }

        public static string GetExcelColumnName(int columnNumber)
        {
            string columnName = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnName;
        }

        public static void ReplaceDatatable(Worksheet worksheet, int rowCount, string datatableName, int rowCursor, int columnCount, int columnCursor = 1, bool showTotals = false, string tableStyle = "TableStyleMedium2")
        {
            // Remove the datatable so we can recreate it
            if (worksheet.ListObjects.Count > 0)
            {

                var allDataTables = worksheet.ListObjects.Cast<object>().Select((t, i) => worksheet.ListObjects[i + 1]).ToList();
                allDataTables.ForEach(list =>
                {                    
                    var matchesName = list.Name == datatableName;
                    if (matchesName)
                    {
                        if (list.DataBodyRange != null)
                        {
                            list.DataBodyRange.ClearFormats();
                        }
                        list.HeaderRowRange.ClearFormats();
                        list.Delete();
                    }
                });
            }

            var dataTable = worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange,
                worksheet.get_Range($"{GetExcelColumnName(columnCursor)}{rowCursor}", $"{GetExcelColumnName(columnCount)}{rowCursor + rowCount - 1}"),
                Type.Missing,
                XlYesNoGuess.xlNo,
                Type.Missing);
            dataTable.Name = datatableName;
            dataTable.ShowTotals = showTotals;
            dataTable.TableStyle = tableStyle;
            
            worksheet.Application.AutoCorrect.AutoFillFormulasInLists = false;
        }

        public static IEnumerable<DataTableContent> GetDataTableContentsRaw(Worksheet worksheet)
        {
            var allDataTables = worksheet.ListObjects.Cast<object>().Select((t, i) => worksheet.ListObjects[i + 1]).ToList();
            return allDataTables.Select(list =>
            {

                var content = new DataTableContent
                {
                    Name = list.Name
                };

                object[,] headers = list.HeaderRowRange.Value;
                foreach (var header in headers)
                {
                    content.Headers.Add(header as string);
                }

                if (list.DataBodyRange == null)
                {
                    return content;
                }

                object[,] body = list.DataBodyRange.Value;

                var dataBodyStartRow = list.DataBodyRange.Row;
                var dataBodyStartColumn = list.DataBodyRange.Column;

                for (var row = 0; row < body.GetLength(0); row++)
                {
                    var rowDictionary = new Dictionary<string, object>();
                    for (var column = 0; column < content.Headers.Count; column++)
                    {
                        var header = content.Headers[column];
                        rowDictionary[header] = worksheet.Cells[dataBodyStartRow + row, dataBodyStartColumn + column].Value;
                    }
                    content.Rows.Add(rowDictionary);
                }

                return content;
            });
        }


    }
}
