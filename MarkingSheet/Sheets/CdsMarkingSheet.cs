using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using MarkingSheet.Model;
using Microsoft.Office.Interop.Excel;
using MarkingSheet.Data;
using System.Windows.Forms;

namespace MarkingSheet.Sheets
{
    internal class CdsMarkingSheet
    {
        const int Name_Column = 1;
        const int Ticker_Column = 2;
        const int Seniority_Column = 3;
        
        private const int OneYear = 4;
        private const int ThreeYear = 5;
        private const int FiveYear = 6;
        private const int SevenYear = 7;
        private const int TenYear = 8;

        const int Sicovam_Column = 9;
        const int ReferenceSicovam_Column = 10;

        private const int IceOneYear = 11;
        private const int IceThreeYear = 12;
        private const int IceFiveYear = 13;
        private const int IceSevenYear = 14;
        private const int IceTenYear = 15;
        private const int IceCurveDate = 16;


        public static void LoadCdsMarkingSheetPositions()
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            // 1. Get Sophis curves
            var sophisCdsCurves = PositionRepository.FetchSophisCdsCurves()                
                .Distinct()
                .ToList();

            if (sophisCdsCurves.Count() == 0)
            {
                DialogResult result = MessageBox.Show(
                    "Timed out waiting for Sophis Data. Please try again.",
                    "Time out Error",
                    MessageBoxButtons.RetryCancel
                );

                if(result == DialogResult.Retry)
                {
                    LoadCdsMarkingSheetPositions();
                }
                return;
            }

            // 2. Get ice curves
            var iceCurves = IceCdsRepository.FetchIceCurves(sophisCdsCurves).Where(x => x != null).ToLookup(x => x.SophisCurveSicovam);

            // 3. Get existing sheet curves and merge with sophis curves
            var curvesAlreadyInSheetLookup = GetCurves(activeSheet)
                .ToImmutableDictionary(existingCurve => existingCurve);

            var sophisCdsCurvesAsMarkingSheetCdsCurve = sophisCdsCurves.Select(sophisCurve =>
            {
                return new MarkingSheetCds()
                {
                    Ticker = sophisCurve.Ticker,
                    Seniority = sophisCurve.Seniority,
                    CurveSicovam = sophisCurve.CurveSicovam,
                    DocClause = sophisCurve.DocClause,
                    Currency = sophisCurve.Currency,
                    SwapSicovam = sophisCurve.SwapSicovam,
                    isIndex = sophisCurve.isIndex
                };
            }); 

            var mergedSophisCuves = sophisCdsCurvesAsMarkingSheetCdsCurve
                .Select(sophisCurve =>
                {
                    var curveAlreadyExistsInSheet = curvesAlreadyInSheetLookup.ContainsKey(sophisCurve);
                    if (curveAlreadyExistsInSheet)
                    {
                        return curvesAlreadyInSheetLookup[sophisCurve];
                    }
                    else
                    { 
                        return sophisCurve; 
                    }
                })
                .OrderBy(curve => curve.isIndex)
                .ThenBy(curve => curve.Ticker)
                .ToList();

            var rowCursor = 1;
            activeSheet.Cells[rowCursor, 1] = "CDS Marks";
            activeSheet.Cells[rowCursor, 1].Font.Bold = true;

            rowCursor = 5;

            Utils.AddDataTable(activeSheet, mergedSophisCuves.Count(), "CDSCurves", rowCursor, 8, 1, false);

            activeSheet.Cells[rowCursor, Name_Column] = "Name";
            activeSheet.Cells[rowCursor, Ticker_Column] = "ICE Ticker";
            activeSheet.Cells[rowCursor, Seniority_Column] = "Seniority";
            activeSheet.Cells[rowCursor, Sicovam_Column] = "Curve Sicovam";
            activeSheet.Cells[rowCursor, ReferenceSicovam_Column] = "Instrument Sicovam";
            activeSheet.Cells[rowCursor, OneYear] = "1Y";
            activeSheet.Cells[rowCursor, ThreeYear] = "3Y";
            activeSheet.Cells[rowCursor, FiveYear] = "5Y";
            activeSheet.Cells[rowCursor, SevenYear] = "7Y";
            activeSheet.Cells[rowCursor, TenYear] = "10Y";

            activeSheet.Cells[rowCursor, IceOneYear] = "ICE 1Y";
            activeSheet.Cells[rowCursor, IceThreeYear] = "ICE 3Y";
            activeSheet.Cells[rowCursor, IceFiveYear] = "ICE 5Y";
            activeSheet.Cells[rowCursor, IceSevenYear] = "ICE 7Y";
            activeSheet.Cells[rowCursor, IceTenYear] = "ICE 10Y";
            activeSheet.Cells[rowCursor, IceCurveDate] = "ICE Curve Date";

            Array.ForEach(new[]
            {
                Name_Column,
                Sicovam_Column,
                Seniority_Column,
                ReferenceSicovam_Column,
                OneYear,
                ThreeYear,
                FiveYear,
                SevenYear,
                TenYear,
                IceCurveDate,
            }, column => activeSheet.Cells[rowCursor, column].Font.Bold = true);
            
            var endRow = rowCursor + 1 + sophisCdsCurves.Count();
            activeSheet.Range[$"{Utils.GetExcelColumnName(OneYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(OneYear)}{endRow}"].NumberFormat = "[Blue]0.000%";
            activeSheet.Range[$"{Utils.GetExcelColumnName(ThreeYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(ThreeYear)}{endRow}"].NumberFormat = "[Blue]0.000%";
            activeSheet.Range[$"{Utils.GetExcelColumnName(FiveYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(FiveYear)}{endRow}"].NumberFormat = "[Blue]0.000%";
            activeSheet.Range[$"{Utils.GetExcelColumnName(SevenYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(SevenYear)}{endRow}"].NumberFormat = "[Blue]0.000%";
            activeSheet.Range[$"{Utils.GetExcelColumnName(TenYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(TenYear)}{endRow}"].NumberFormat = "[Blue]0.000%";

            activeSheet.Range[$"{Utils.GetExcelColumnName(IceOneYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(IceOneYear)}{endRow}"].NumberFormat = "0.000%";
            activeSheet.Range[$"{Utils.GetExcelColumnName(IceThreeYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(IceThreeYear)}{endRow}"].NumberFormat = "0.000%";
            activeSheet.Range[$"{Utils.GetExcelColumnName(IceFiveYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(IceFiveYear)}{endRow}"].NumberFormat = "0.000%";
            activeSheet.Range[$"{Utils.GetExcelColumnName(IceSevenYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(IceSevenYear)}{endRow}"].NumberFormat = "0.000%";
            activeSheet.Range[$"{Utils.GetExcelColumnName(IceTenYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(IceTenYear)}{endRow}"].NumberFormat = "0.000%";

            mergedSophisCuves.ForEach(curve => {

                rowCursor++;

                object[,] payloadToApply = new object[1, 17];
                payloadToApply[0, Ticker_Column - 1] = curve.Ticker;
                payloadToApply[0, Seniority_Column - 1] = curve.Seniority;
                payloadToApply[0, Sicovam_Column - 1] = curve.CurveSicovam;
                payloadToApply[0, ReferenceSicovam_Column - 1] = curve.SwapSicovam;

                payloadToApply[0, OneYear - 1] = curve.OneYear;
                payloadToApply[0, ThreeYear - 1] = curve.ThreeYear;
                payloadToApply[0, FiveYear - 1] = curve.FiveYear;
                payloadToApply[0, SevenYear - 1] = curve.SevenYear;
                payloadToApply[0, TenYear - 1] = curve.TenYear;

                var iceCurve = iceCurves[curve.CurveSicovam].FirstOrDefault();
                if (iceCurve != null)
                {
                    payloadToApply[0, Name_Column - 1] = iceCurve.Name;
                    payloadToApply[0, IceOneYear - 1] = iceCurve.OneYear;
                    payloadToApply[0, IceThreeYear - 1] = iceCurve.ThreeYear;
                    payloadToApply[0, IceFiveYear - 1] = iceCurve.FiveYear;
                    payloadToApply[0, IceSevenYear - 1] = iceCurve.SevenYear;
                    payloadToApply[0, IceTenYear - 1] = iceCurve.TenYear;
                    payloadToApply[0, IceCurveDate - 1] = iceCurve.IceCurveDate;
                }

                activeSheet.Range[$"A{rowCursor}", $"{Utils.GetExcelColumnName(16)}{rowCursor}"].Formula = payloadToApply;

            });

            activeSheet.Columns.AutoFit();
        }

        public static List<MarkingSheetCds> GetCurves(Worksheet activeSheet)
        {
            var currentTable = Utils.GetDataTableContentsRaw(activeSheet);

            var existingCurves = new List<MarkingSheetCds>();

            if (currentTable.Count() > 0 && currentTable.First().Name == "CDSCurves")
            {
                var existingTable = currentTable.First();

                existingTable.Rows.ForEach(row =>
                {
                    var ticker = row["ICE Ticker"].ToString();
                    var name = row["Name"]?.ToString() ?? "";
                    var seniority = row["Seniority"].ToString();
                    var curveSicovam = row["Curve Sicovam"].ToString();
                    var swapSicovam = row["Instrument Sicovam"].ToString();
                    double? oneYear = null;
                    if (row["1Y"] != null)
                    {
                        oneYear = Convert.ToDouble(row["1Y"].ToString());
                    }

                    double? threeYear = null;
                    if (row["3Y"] != null) {
                        threeYear =Convert.ToDouble(row["3Y"].ToString());
                    }

                    double? fiveYear = null;
                    if (row["5Y"] != null) {
                        fiveYear = Convert.ToDouble(row["5Y"].ToString());
                    }

                    double? sevenYear = null;
                    if (row["7Y"] != null) {
                        sevenYear = Convert.ToDouble(row["7Y"].ToString());
                    }

                    double? tenYear = null;
                    if (row["10Y"] != null)
                    {
                        tenYear = Convert.ToDouble(row["10Y"].ToString());
                    }
                    
                    existingCurves.Add(new MarkingSheetCds
                    {
                        Ticker = ticker,
                        Seniority = seniority,
                        CurveSicovam = Convert.ToInt32(curveSicovam),
                        SwapSicovam = Convert.ToInt32(swapSicovam),
                        OneYear = oneYear,
                        ThreeYear = threeYear,
                        FiveYear = fiveYear,
                        SevenYear = sevenYear,
                        TenYear = tenYear,
                        isIndex = name.Contains("iTraxx")
                    });

                });

            }

            return existingCurves;
        }
    }
}
