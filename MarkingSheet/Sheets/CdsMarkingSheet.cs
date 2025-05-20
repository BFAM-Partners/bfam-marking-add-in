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
        
        private const int OneYear = 2;
        private const int ThreeYear = 3;
        private const int FiveYear = 4;
        private const int SevenYear = 5;
        private const int TenYear = 6;        

        private const int IceOneYear = 7;
        private const int IceThreeYear = 8;
        private const int IceFiveYear = 9;
        private const int IceSevenYear = 10;
        private const int IceTenYear = 11;
        private const int SevenMinusFiveYear = 12;
        private const int TenMinusFiveYear = 13;
        private const int FiveMinusThreeYear = 14;
        private const int IceCurveDate = 15;

        const int Ticker_Column = 16;
        const int Seniority_Column = 17;
        private const int Currency = 18;
        private const int DocumentClause = 19;        

        const int Sicovam_Column = 20;
        const int ReferenceSicovam_Column = 21;


        public static void LoadCdsMarkingSheetPositions()
        {
            Worksheet activeSheet = (Worksheet) Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            var sophisCdsCurves = PositionRepository.FetchSophisCdsCurves()                
                .Distinct()
                .ToList();

            if (sophisCdsCurves.Count() == 0)
            {
                DialogResult result = MessageBox.Show(
                    "Timed out waiting for Sophis Data. Please try again in five minutes.",
                    "Time out Error",
                    MessageBoxButtons.RetryCancel
                );

                if(result == DialogResult.Retry)
                {
                    LoadCdsMarkingSheetPositions();
                }
                return;
            }


            // Curves from Marking Sheet
            var curvesAlreadyInSheetLookup = GetCurves(activeSheet, tableName: "CDSCurves").ToImmutableDictionary(existingCurve => existingCurve);

            // Curves from Tracker
            var trackerCurvesInSheet = GetCurves(activeSheet, tableName: "CDSCurveTracker").Where(curve => curve.Ticker != null && curve.Ticker != "").ToList();

            var droppedRows = curvesAlreadyInSheetLookup.Values.Where(alreadyInSheet => sophisCdsCurves.Where(sophisCurve => sophisCurve.CurveSicovam == alreadyInSheet.CurveSicovam).Count() == 0)
                .Where(droppedRow => trackerCurvesInSheet.Where(trackerCurve => trackerCurve.Ticker == droppedRow.Ticker).Count() == 0); // Make sure its not already in Tracker

            trackerCurvesInSheet.AddRange(droppedRows); // Only where they dont exist

            var sophisCdsCurvesAsMarkingSheetCdsCurve = mapSophisPositionsToMarkingSheetCdsCurves(sophisCdsCurves, curvesAlreadyInSheetLookup);

            IceCdsRepository.EnrichMarkingSheetCdsCurveWithIceCds(sophisCdsCurvesAsMarkingSheetCdsCurve);

            // Tracker
            IceCdsRepository.EnrichMarkingSheetCdsCurveWithIceCds(trackerCurvesInSheet);

            // Clear all borders            
            clearBorders(activeSheet, new int[]{ IceOneYear, SevenMinusFiveYear, IceCurveDate, Ticker_Column, Sicovam_Column });
            
            var rowCursor = 1;
            activeSheet.Cells[rowCursor, 1] = "CDS Marks";
            activeSheet.Cells[rowCursor, 1].Font.Bold = true;

            rowCursor = 4;

            activeSheet.Range[activeSheet.Cells[rowCursor-1, Ticker_Column], activeSheet.Cells[rowCursor-1, DocumentClause]].Merge();
            activeSheet.Range[activeSheet.Cells[rowCursor - 1, Ticker_Column], activeSheet.Cells[rowCursor - 1, DocumentClause]].Value = "Required Fields";
            activeSheet.Range[activeSheet.Cells[rowCursor - 1, Ticker_Column], activeSheet.Cells[rowCursor - 1, DocumentClause]].Font.Bold = true;
            activeSheet.Range[activeSheet.Cells[rowCursor - 1, Ticker_Column], activeSheet.Cells[rowCursor - 1, DocumentClause]].HorizontalAlignment = XlHAlign.xlHAlignCenter;

            rowCursor = CurveTable(activeSheet,
                curves: trackerCurvesInSheet,
                rowCursor: rowCursor,
                tableName: "CDSCurveTracker",
                tableStyle: "TableStyleMedium4",
                includeSicovam: false
                );

            rowCursor += 4;

            rowCursor = CurveTable(activeSheet,
                curves: sophisCdsCurvesAsMarkingSheetCdsCurve,
                rowCursor: rowCursor,
                tableName: "CDSCurves",
                tableStyle: "TableStyleMedium2",
                includeSicovam: true
                );


            activeSheet.Columns.AutoFit();
        }

        private static void clearBorders(Worksheet activeSheet, int[] columns)
        {
            foreach (var column in columns)
            {
                activeSheet.Range[activeSheet.Cells[1, column], activeSheet.Cells[1 + 250, column]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlLineStyleNone;
            }            
        }

        private static int CurveTable(Worksheet activeSheet, List<MarkingSheetCds> curves, int rowCursor, string tableName, string tableStyle, bool includeSicovam)
        {
            var numberOfColumns = 18;
            if (includeSicovam)
            {
                numberOfColumns = 20;
            }

            Utils.ReplaceDatatable(activeSheet, Math.Max(curves.Count(), 1), tableName, rowCursor, numberOfColumns, 1, false, tableStyle: tableStyle);

            activeSheet.Cells[rowCursor, Name_Column] = "Name";
            activeSheet.Cells[rowCursor, Ticker_Column] = "ICE Ticker";
            activeSheet.Cells[rowCursor, Seniority_Column] = "Seniority";

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

            activeSheet.Cells[rowCursor, SevenMinusFiveYear] = "ICE 7Y-5Y";
            activeSheet.Cells[rowCursor, TenMinusFiveYear] = "ICE 10T-5Y";
            activeSheet.Cells[rowCursor, FiveMinusThreeYear] = "ICE 5Y-3Y";

            activeSheet.Cells[rowCursor, Currency] = "Currency";
            activeSheet.Cells[rowCursor, DocumentClause] = "Document Clause";

            if (includeSicovam)
            {
                activeSheet.Cells[rowCursor, Sicovam_Column] = "Curve Sicovam";
                activeSheet.Cells[rowCursor, ReferenceSicovam_Column] = "Instrument Sicovam";
            }

            var endRow = rowCursor + 1 + curves.Count();
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

            activeSheet.Range[$"{Utils.GetExcelColumnName(SevenMinusFiveYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(SevenMinusFiveYear)}{endRow}"].NumberFormat = "0.000%";
            activeSheet.Range[$"{Utils.GetExcelColumnName(TenMinusFiveYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(TenMinusFiveYear)}{endRow}"].NumberFormat = "0.000%";
            activeSheet.Range[$"{Utils.GetExcelColumnName(FiveMinusThreeYear)}{rowCursor + 1}", $"{Utils.GetExcelColumnName(FiveMinusThreeYear)}{endRow}"].NumberFormat = "0.000%";

            activeSheet.Range[activeSheet.Cells[rowCursor, IceOneYear], activeSheet.Cells[rowCursor + curves.Count(), IceOneYear]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            activeSheet.Range[activeSheet.Cells[rowCursor, SevenMinusFiveYear], activeSheet.Cells[rowCursor + curves.Count(), SevenMinusFiveYear]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            activeSheet.Range[activeSheet.Cells[rowCursor, IceCurveDate], activeSheet.Cells[rowCursor + curves.Count(), IceCurveDate]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            activeSheet.Range[activeSheet.Cells[rowCursor, Ticker_Column], activeSheet.Cells[rowCursor + curves.Count(), Ticker_Column]].Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;

            curves.ForEach(curve => {

                rowCursor++;

                object[,] payloadToApply = new object[1, numberOfColumns + 1];
                payloadToApply[0, Ticker_Column - 1] = curve.Ticker;
                payloadToApply[0, Seniority_Column - 1] = curve.Seniority;

                payloadToApply[0, OneYear - 1] = curve.OneYear;
                payloadToApply[0, ThreeYear - 1] = curve.ThreeYear;
                payloadToApply[0, FiveYear - 1] = curve.FiveYear;
                payloadToApply[0, SevenYear - 1] = curve.SevenYear;
                payloadToApply[0, TenYear - 1] = curve.TenYear;

                var iceCurve = curve.iceCdsCurve;
                if (iceCurve != null)
                {
                    payloadToApply[0, Name_Column - 1] = iceCurve.Name;
                    payloadToApply[0, IceOneYear - 1] = iceCurve.OneYear;
                    payloadToApply[0, IceThreeYear - 1] = iceCurve.ThreeYear;
                    payloadToApply[0, IceFiveYear - 1] = iceCurve.FiveYear;
                    payloadToApply[0, IceSevenYear - 1] = iceCurve.SevenYear;
                    payloadToApply[0, IceTenYear - 1] = iceCurve.TenYear;
                    payloadToApply[0, IceCurveDate - 1] = iceCurve.IceCurveDate;
                    payloadToApply[0, SevenMinusFiveYear - 1] = "=[@[ICE 7Y]]-[@[ICE 5Y]]";
                    payloadToApply[0, TenMinusFiveYear - 1] = "=[@[ICE 10Y]]-[@[ICE 5Y]]";
                    payloadToApply[0, FiveMinusThreeYear - 1] = "=[@[ICE 5Y]]-[@[ICE 3Y]]";
                }

                payloadToApply[0, Currency - 1] = curve.Currency;
                payloadToApply[0, DocumentClause - 1] = curve.DocClause;

                if (includeSicovam)
                {
                    payloadToApply[0, Sicovam_Column - 1] = curve.CurveSicovam;
                    payloadToApply[0, ReferenceSicovam_Column - 1] = curve.SwapSicovam;
                }

                activeSheet.Range[$"A{rowCursor}", $"{Utils.GetExcelColumnName(numberOfColumns + 1)}{rowCursor}"].Formula = payloadToApply;

            });
            return rowCursor;
        }

        private static List<MarkingSheetCds> mapSophisPositionsToMarkingSheetCdsCurves(List<SophisCdsCurve> sophisCdsCurves, ImmutableDictionary<MarkingSheetCds, MarkingSheetCds> curvesAlreadyInSheetLookup)
        {
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

            return sophisCdsCurvesAsMarkingSheetCdsCurve
                .Select(sophisCurve =>
                {
                    var curveAlreadyExistsInSheet = curvesAlreadyInSheetLookup.ContainsKey(sophisCurve);
                    if (curveAlreadyExistsInSheet)
                    {
                        var sheetCurve = curvesAlreadyInSheetLookup[sophisCurve];
                        sophisCurve.OneYear = sheetCurve.OneYear;
                        sophisCurve.ThreeYear = sheetCurve.ThreeYear;
                        sophisCurve.FiveYear = sheetCurve.FiveYear;
                        sophisCurve.SevenYear = sheetCurve.SevenYear;
                        sophisCurve.TenYear = sheetCurve.TenYear;

                        return sophisCurve;
                    }
                    else
                    {
                        return sophisCurve;
                    }
                })
                .OrderBy(curve => curve.isIndex)
                .ThenBy(curve => curve.Ticker)
                .ToList();
        }

        public static List<MarkingSheetCds> GetCurves(Worksheet activeSheet, string tableName)
        {
            var currentTable = Utils.GetDataTableContentsRaw(activeSheet);

            var existingCurves = new List<MarkingSheetCds>();

            foreach (var table in currentTable) {

                if (table.Name == tableName)
                {
                    var existingTable = table;

                    existingTable.Rows.ForEach(row =>
                    {
                        var ticker = row["ICE Ticker"]?.ToString();
                        var name = row["Name"]?.ToString() ?? "";
                        var seniority = row["Seniority"]?.ToString();
                        
                        var curveSicovam = 0;
                        if (row.ContainsKey("Curve Sicovam"))
                        {
                            curveSicovam = Convert.ToInt32(row["Curve Sicovam"]?.ToString());
                        }

                        var swapSicovam = 0;
                        if (row.ContainsKey("Instrument Sicovam"))
                        {
                            swapSicovam = Convert.ToInt32(row["Instrument Sicovam"]?.ToString());
                        }

                        
                        double? oneYear = null;
                        if (row["1Y"] != null)
                        {
                            oneYear = Convert.ToDouble(row["1Y"]?.ToString());
                        }

                        double? threeYear = null;
                        if (row["3Y"] != null)
                        {
                            threeYear = Convert.ToDouble(row["3Y"].ToString());
                        }

                        double? fiveYear = null;
                        if (row["5Y"] != null)
                        {
                            fiveYear = Convert.ToDouble(row["5Y"].ToString());
                        }

                        double? sevenYear = null;
                        if (row["7Y"] != null)
                        {
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
                            CurveSicovam = curveSicovam,
                            SwapSicovam = swapSicovam,
                            OneYear = oneYear,
                            ThreeYear = threeYear,
                            FiveYear = fiveYear,
                            SevenYear = sevenYear,
                            TenYear = tenYear,
                            isIndex = ticker != null && ticker.Contains("ITX"),
                            Currency = row["Currency"]?.ToString(),
                            DocClause = row["Document Clause"]?.ToString()
                        });

                    });
                }
            }
            return existingCurves;
        }        
    }
}
