using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;

namespace MarkingSheet
{

    internal class BondMarkingSheet
    {
        const int Isin_Column = 1;
        const int Name_Column = 2;
        const int Sicovam_Column = 3;
        const int MarkSource_Column = 4;
        const int MarkToUpload_Column = 5;
        const int CBBT_Column = 6;
        const int BVAL_Column = 7;
        const int Manual_Column = 8;

        private static Dictionary<string, string> GetExistingManualPrices(IEnumerable<Utils.DataTableContent> tables, string tableName)
        {
            var existingManualPrices = new Dictionary<string, string>();

            foreach (Utils.DataTableContent table in tables) {

                if (table.Name == tableName)
                {
                    table.Rows.ForEach(row =>
                    {
                        if (row["Isin"] == null)
                        {
                           return;
                        }
                        var isin = row["Isin"].ToString();
                        //var mark = row["Mark to Upload"].ToString();
                        //var cbbt = row["CBBT"].ToString();
                        //var bval = row["BVAL"].ToString();
                        var manual = row["Manual"]?.ToString().Trim();

                        existingManualPrices.Add(isin, manual);
                        //if (manual != null && manual.Trim().Length > 0)
                        //{
                        //    existingManualPrices.Add(isin, manual);
                        //}
                    });

                }
            }

            return existingManualPrices;
        }

        public static void LoadBondMarkingSheetPositions()
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet); 

            var existingTableContents = Utils.GetDataTableContentsRaw(activeSheet);

            var existingManualPricesForPositions = GetExistingManualPrices(existingTableContents, "Positions");
            var existingManualPricesForTracker = GetExistingManualPrices(existingTableContents, "BondTracker");

            activeSheet.Cells[1, 1].Font.Bold = true;

            // Check if any rows already exist, if so grab the marks
            var bonds = PositionRepository.FetchBonds().Where(bond => bond.Quantity != 0.0).ToList();

            // Add exclusions                   
            bonds = bonds.Where(bond => ExclusionList.excludedIsins().Contains(bond.Isin) == false).ToList();

            bonds = bonds.Select(bond =>
            {
                if (existingManualPricesForPositions.ContainsKey(bond.Isin))
                {
                    bond.ManualPrice = existingManualPricesForPositions[bond.Isin];
                }
                return bond;
            }).OrderByDescending(bond => bond.Isin).ToList();
            
            foreach (var entry in existingManualPricesForPositions)
            {
                var foundBond = bonds.Where(bond => bond.Isin == entry.Key).Count() != 0;

                if (!foundBond)
                {
                    if (!existingManualPricesForTracker.ContainsKey(entry.Key))
                    {
                        existingManualPricesForTracker[entry.Key] = entry.Value;
                    }
                }
            }

            var rowCursor = 1;
            activeSheet.Cells[rowCursor, 1] = "Bond Marks";

            rowCursor = 4;            

            Utils.ReplaceDatatable(activeSheet, Math.Max(existingManualPricesForTracker.Count, 1), "BondTracker", rowCursor, 5, 1, showTotals: false, tableStyle: "TableStyleMedium4");

            activeSheet.Cells[rowCursor, Isin_Column] = "Isin";
            activeSheet.Cells[rowCursor, Name_Column] = "Name";
            activeSheet.Cells[rowCursor, CBBT_Column - 3] = "CBBT";
            activeSheet.Cells[rowCursor, BVAL_Column - 3] = "BVAL";
            activeSheet.Cells[rowCursor, Manual_Column - 3] = "Manual";

            existingManualPricesForTracker.ToList().ForEach(existingManualPrice => {

                rowCursor++;

                var isin = existingManualPrice.Key;
                var manualPrice = existingManualPrice.Value;

                object[,] payloadToApply = new object[1, 8];
                var cleanIsin = isin.Replace("\"", "");
                payloadToApply[0, Isin_Column - 1] = cleanIsin;
                payloadToApply[0, Name_Column - 1] = $@"=BDP([@Isin] & "" Corp"", ""SECURITY_NAME"")"; ;                
                
                //activeSheet.Cells[rowCursor, 4] = bond.Quantity;                                
                payloadToApply[0, CBBT_Column - 4] = $@"=BDP([@Isin] & ""@CBBT Corp"", ""PX_LAST"")";
                payloadToApply[0, BVAL_Column - 4] = $@"=BDP([@Isin] & ""@BVAL Corp"", ""PX_LAST"")";
                payloadToApply[0, Manual_Column - 4] = manualPrice;

                activeSheet.Range[$"A{rowCursor}", $"{Utils.GetExcelColumnName(8)}{rowCursor}"].Formula = payloadToApply;
            });

            rowCursor += 4;

            Utils.ReplaceDatatable(activeSheet, bonds.Count, "Positions", rowCursor, 7, 1, false);


            activeSheet.Cells[rowCursor, Isin_Column] = "Isin";
            activeSheet.Cells[rowCursor, Name_Column] = "Name";
            activeSheet.Cells[rowCursor, Sicovam_Column] = "Sicovam";
            activeSheet.Cells[rowCursor, MarkSource_Column] = "Mark source";
            activeSheet.Cells[rowCursor, MarkToUpload_Column] = "Mark to Upload";
            activeSheet.Cells[rowCursor, CBBT_Column] = "CBBT";
            activeSheet.Cells[rowCursor, BVAL_Column] = "BVAL";
            activeSheet.Cells[rowCursor, Manual_Column] = "Manual";


            activeSheet.Cells[rowCursor, Isin_Column].Font.Bold = true;
            activeSheet.Cells[rowCursor, Name_Column].Font.Bold = true;
            activeSheet.Cells[rowCursor, Sicovam_Column].Font.Bold = true;
            activeSheet.Cells[rowCursor, MarkSource_Column].Font.Bold = true;
            activeSheet.Cells[rowCursor, MarkToUpload_Column].Font.Bold = true;
            activeSheet.Cells[rowCursor, CBBT_Column].Font.Bold = true;
            activeSheet.Cells[rowCursor, BVAL_Column].Font.Bold = true;
            activeSheet.Cells[rowCursor, Manual_Column].Font.Bold = true;

            bonds.ForEach(bond => {

                rowCursor++;


                object[,] payloadToApply = new object[1, 8];
                var cleanIsin = bond.Isin.Replace("\"", "");
                payloadToApply[0, Isin_Column - 1] = bond.Isin;
                payloadToApply[0, Name_Column - 1] = $@"=BDP(""{cleanIsin} Corp"", ""SECURITY_NAME"")"; ;
                payloadToApply[0, Sicovam_Column - 1] = bond.Sicovam;


                payloadToApply[0, MarkSource_Column - 1] = $@"=IF(ISNUMBER([@Manual]), ""Manual"", IF(ISNUMBER([@CBBT]), ""CBBT"", ""BVAL""))";
                //activeSheet.Cells[rowCursor, 4] = bond.Quantity;                
                payloadToApply[0, MarkToUpload_Column - 1] = $@"=IF(ISNUMBER([@Manual]), [@Manual], IF(ISNUMBER([@CBBT]), [@CBBT], [@BVAL]))";
                payloadToApply[0, CBBT_Column - 1] = $@"=BDP(""{cleanIsin}@CBBT Corp"", ""PX_LAST"")";
                payloadToApply[0, BVAL_Column - 1] = $@"=BDP(""{cleanIsin}@BVAL Corp"", ""PX_LAST"")";
                payloadToApply[0, Manual_Column - 1] = bond.ManualPrice;

                activeSheet.Range[$"A{rowCursor}", $"{Utils.GetExcelColumnName(8)}{rowCursor}"].Formula = payloadToApply;
            });

            activeSheet.Columns.AutoFit();
        }

    }
}
