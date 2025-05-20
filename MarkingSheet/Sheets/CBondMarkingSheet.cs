using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;

namespace MarkingSheet
{
    internal class CBondMarkingSheet
    {
        private static Range xlRange;
        private static Worksheet activeSheet;

        public enum Columns
        {
            Isin = 1,
            UnderlyingReference,
            Name,
            Portfolio,
            Country,
            BVALDoDDiff,
      
            Separator1,
            SophisPriceEOD,
            BVALEOD,
            BvalEODDiff,
            TotalEODPNL,
            BondEODPNL,
            AscotEODPNL,
            Separator2, 
            SophisPrice,
            BVAL,
            BvalDiff,
            TotalPNL,
            BondPNL,
            AscotPNL,
            Separator3,
            
            Volatility,
            SophisFloor,
            Parity,
            Par,
            ConvRatio,
            Premium,
            Spread,
            Borrow,
            Last,
            Theoretical,
            LastEOD,
            TheoreticalEOD,
            SophisPriceSource,
            SophisMV,
            BVALMV,
            SophisEODMV,
            BVALEODMV,
            BondPosition,
            AscotPosition,
            BondNotional,
            BondNotionalEOD,
            AscotNotional,
            AscotNotionalEOD,
            FX,
            FX_EOD,
            BondCurrency,
            UnderlyingCurrency,
            ImpliedVol,
            NukedDiff,
            NukedMid,
            BestBidNuked,
            BestAskNuked,
            BestBidIssuer,
            BestBid,
            BestBidRef,
            BestAskIssuer,
            BestAsk,
            BestAskRef,
            UnderlyingPrice,
            UnderlyingPriceInBondCurrency,
            Delta,
            Sicovam
        }

        static Dictionary<int, string> columnNames = new Dictionary<int, string>
        {
            {(int)Columns.Isin, "Isin"},
            {(int)Columns.UnderlyingReference, "Underlying"},
            {(int)Columns.Name, "Name"},
            {(int)Columns.Portfolio, "Portfolio"},
            {(int)Columns.Country, "Country"},
            {(int)Columns.Last, "Last"},
            {(int)Columns.Theoretical, "Theoretical"},
            {(int)Columns.LastEOD, "EOD Last"},
            {(int)Columns.TheoreticalEOD, "EOD Theoretical"},
            {(int)Columns.SophisPrice, "Sophis"},
            {(int)Columns.BVALDoDDiff, "BVAL DoD Diff"},
            {(int)Columns.SophisPriceEOD, "Sophis EOD"},
            {(int)Columns.SophisPriceSource, "Sophis Price Source"},
            {(int)Columns.Separator1, "..."},
            {(int)Columns.Separator2, "..."},
            {(int)Columns.Separator3, "..."},
            {(int)Columns.BVAL, "BVAL"},
            {(int)Columns.BVALEOD, "BVAL EOD"},
            {(int)Columns.NukedDiff, "Nuked Diff"},
            {(int)Columns.NukedMid, "Nuked Mid"},
            {(int)Columns.BvalDiff, "BVAL Diff"},
            {(int)Columns.BvalEODDiff, "BVAL EOD Diff"},
            {(int)Columns.BondPNL, "Bond PnL"},
            {(int)Columns.AscotPNL, "Ascot PnL"},
            {(int)Columns.TotalPNL, "PnL"},
            {(int)Columns.BondEODPNL, "Bond EOD PnL"},
            {(int)Columns.AscotEODPNL, "Ascot EOD PnL"},
            {(int)Columns.TotalEODPNL, "EOD PnL"},
            {(int)Columns.BVALMV, "BVAL MV"},
            {(int)Columns.SophisMV, "Sophis MV"},
            {(int)Columns.BVALEODMV, "BVAL EOD MV"},
            {(int)Columns.SophisEODMV, "Sophis EOD MV"},
            {(int)Columns.BondPosition, "Bond Position"},
            {(int)Columns.AscotPosition, "Ascot Position"},
            {(int)Columns.BondNotional, "Bond Notional"},
            {(int)Columns.AscotNotional, "Ascot Notional"},
            {(int)Columns.BondNotionalEOD, "Bond Notional EOD"},
            {(int)Columns.AscotNotionalEOD, "Ascot Notional EOD"},
            {(int)Columns.SophisFloor, "Bond Floor"}, //TODO: change to Sophis floor when sourced from sophis
            {(int)Columns.Par, "Par"},
            {(int)Columns.Parity, "Parity"},
            {(int)Columns.ConvRatio, "Conv Ratio"},
            {(int)Columns.Premium, "Premium %"},
            {(int)Columns.FX, "FX Rate"},
            {(int)Columns.FX_EOD, "FX EOD"},
            {(int)Columns.BondCurrency, "Bond Currency"},
            {(int)Columns.UnderlyingCurrency, "Underlying Currency"},
            {(int)Columns.Volatility, "Sophis Vol"},
            {(int)Columns.ImpliedVol, "Best Bid Implied Vol"},
            {(int)Columns.BestBidNuked, "Best Bid Nuked"},
            {(int)Columns.BestAskNuked, "Best Ask Nuked"},
            {(int)Columns.BestBidIssuer, "Best Bid Issuer"},
            {(int)Columns.BestBid, "Best Bid"},
            {(int)Columns.BestBidRef, "Best Bid Ref"},
            {(int)Columns.BestAskIssuer, "Best Ask Issuer"},
            {(int)Columns.BestAsk, "Best Ask"},
            {(int)Columns.BestAskRef, "Best Ask Ref"},
            {(int)Columns.UnderlyingPrice, "Sophis Underlying Price"},
            {(int)Columns.UnderlyingPriceInBondCurrency, "Underlying in Bond Curr"},
            {(int)Columns.Delta, "Sophis $ Delta"},
            {(int)Columns.Spread, "Spread"},
            {(int)Columns.Borrow, "Borrow"},
            {(int)Columns.Sicovam, "Sicovam"},
        };

        static int numOfColumns = Enum.GetNames(typeof(Columns)).Length;

        static List<string> defaultedCBs = new List<string> { "XS1580153408", "XS0911050390", "XS0880097620", "XS1057356773", "XS0683220650", "XS0979033106" };
        static List<string> markedToLast = new List<string> { "XS2771889610", "XS2089158609", "XS2466214629", "XS2785324703" };

        public static void ApplyBuiltInConditionalFormatting(int lastRow)
        {
            Worksheet activeSheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            Range bvalDiffRange = activeSheet.Range[activeSheet.Cells[6, (int)Columns.BvalDiff], activeSheet.Cells[lastRow, (int)Columns.BvalDiff]];
            Range bvalEODDiffRange = activeSheet.Range[activeSheet.Cells[6, (int)Columns.BvalEODDiff], activeSheet.Cells[lastRow, (int)Columns.BvalEODDiff]];
            Range nukedDiffRange = activeSheet.Range[activeSheet.Cells[6, (int)Columns.NukedDiff], activeSheet.Cells[lastRow, (int)Columns.NukedDiff]];

            ApplyFormatConditions(bvalDiffRange);
            ApplyFormatConditions(bvalEODDiffRange);
            ApplyFormatConditions(nukedDiffRange);
        }

        private static void ApplyFormatConditions(Range range)
        {
            int colorRed = ColorTranslator.ToOle(ColorTranslator.FromHtml("#F8696B"));
            int colorOrange = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFA500"));
            int colorYellow = ColorTranslator.ToOle(ColorTranslator.FromHtml("#FFEB84"));
            int colorGreen = ColorTranslator.ToOle(ColorTranslator.FromHtml("#63BE7B"));

            range.FormatConditions.Delete();

            // Condition 1: Absolute Difference > 10 (Red)
            string formula1 = $"=ABS({range.Address[false, false]})>10";
            FormatCondition cond1 = (FormatCondition)range.FormatConditions.Add(XlFormatConditionType.xlExpression, Formula1: formula1);
            cond1.Interior.Color = colorRed;

            // Condition 2: Absolute Difference > 0.3 and <= 1 (Yellow)
            FormatCondition cond2Negative = (FormatCondition)range.FormatConditions.Add(
                XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlBetween, "=-0.3", "=-1");
            cond2Negative.Interior.Color = colorYellow;
            FormatCondition cond2Positive = (FormatCondition)range.FormatConditions.Add(
                XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlBetween, "=0.3", "=1");
            cond2Positive.Interior.Color = colorYellow;


            // Condition 3: Absolute Difference <= 0.3 (Green)
            string formula3 = $"=ABS({range.Address[false, false]})<=0.3";
            FormatCondition cond3 = (FormatCondition)range.FormatConditions.Add(XlFormatConditionType.xlExpression, Formula1: formula3);
            cond3.Interior.Color = colorGreen;

            // Condition 4: Absolute Difference > 1 and <= 10 (Orange)
            FormatCondition cond4Negative = (FormatCondition)range.FormatConditions.Add(
                XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlBetween, "=-1", "=-10");
            cond4Negative.Interior.Color = colorOrange;
            FormatCondition cond4Positive = (FormatCondition)range.FormatConditions.Add(
                XlFormatConditionType.xlCellValue, XlFormatConditionOperator.xlBetween, "=1", "=10");
            cond4Positive.Interior.Color = colorOrange;
        }

        public static void LoadCBondMarkingSheetPositions()
        {
            activeSheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);

            var convertibleBonds = PositionRepository.FetchConvertibles()
                .GroupBy(cbond => cbond.Isin)
                .Select(group =>
                {
                    double? ascotPosition = 0.0;
                    double? bondPosition = 0.0;
                    foreach (var instrument in group)
                    {
                        var instrumentType = instrument.InstrumentType;
                        if(instrumentType == "ascot")
                        {
                            ascotPosition = ascotPosition + instrument.PositionUSD;
                        }
                        else if(instrumentType == "convertibleBond")
                        {
                            bondPosition = bondPosition + instrument.PositionUSD;
                        }
                        else if (instrumentType == "swap") // cbSwaps
                        {
                            bondPosition = bondPosition + instrument.PositionUSD;
                        }
                    }
                    return new ConvertibleBond
                    {
                        Isin = group.Key,
                        BondPosition = bondPosition,
                        AscotPosition = ascotPosition,
                        Name = group.First().Name,
                        Portfolio = group.First().Portfolio,
                        UnderlyingReference = group.First().UnderlyingReference,
                        Sicovam = group.First().Sicovam,
                        Volatility = group.First().Volatility,
                        Spread = group.First().Spread,
                        Theoretical = group.First().Theoretical,
                        Last = group.First().Last,
                        HistoricalTheoretical = group.First().HistoricalTheoretical,
                        HistoricalLast = group.First().HistoricalLast,
                        UnderlyingPrice = group.First().UnderlyingPrice,
                        Delta = group.First().Delta,
                        Borrow = group.First().Borrow,
                        Currency = group.First().Currency,
                        BestBidIssuer = group.First().BestBidIssuer,
                        BestBidBid = group.First().BestBidBid,
                        BestBidRef = group.First().BestBidRef,
                        BestAskIssuer = group.First().BestAskIssuer,
                        BestAskAsk = group.First().BestAskAsk,
                        BestAskRef = group.First().BestAskRef,
                        BestBidNuked = group.First().BestBidNuked,
                        BestAskNuked = group.First().BestAskNuked,
                    };
                })
                .Where(cbond => !defaultedCBs.Contains(cbond.Isin))
                .OrderBy(cbond => cbond.Portfolio)
                .ToList();

            //Clean sheet
            try
            {
                activeSheet.Cells.UnMerge();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error while trying to unmerge cells", e.Message);
            }
            try
            {
                activeSheet.Cells.Clear();
            }
            catch (Exception e)
            {
                Console.WriteLine("Error while trying to clear cells", e.Message);
                activeSheet.Range["A4:AAA200"].Value = "";
            }
            
            if (activeSheet.AutoFilterMode)
            {
                activeSheet.AutoFilterMode = false;
            }
            activeSheet.Parent.Windows[1].Zoom = 100;

            activeSheet.Cells[1, 1] = "Convertible Bond Marks";
            activeSheet.Cells[1, 1].Font.Bold = true;

            activeSheet.Cells[1, 3] = "Last updated at";
            activeSheet.Cells[1, 3].Font.Bold = true;
            activeSheet.Cells[1, 3].HorizontalAlignment = XlHAlign.xlHAlignRight;
            activeSheet.Cells[2, 3] = DateTime.Now.ToString("dd/MM/yyyy HH:mm");

            var rowCursor = 5;

            int sectionHeaderRow = rowCursor - 1;

            // Set up section headers and merge cells 
            Range mergeEOD = activeSheet.Range[activeSheet.Cells[sectionHeaderRow, (int)Columns.SophisPriceEOD], activeSheet.Cells[sectionHeaderRow, (int)Columns.AscotEODPNL]];
            Range mergeCurrent = activeSheet.Range[activeSheet.Cells[sectionHeaderRow, (int)Columns.SophisPrice], activeSheet.Cells[sectionHeaderRow, (int)Columns.AscotPNL]];
            Range mergeMarkComparison = activeSheet.Range[activeSheet.Cells[sectionHeaderRow, (int)Columns.NukedDiff], activeSheet.Cells[sectionHeaderRow, (int)Columns.ImpliedVol]];
            Range mergeQuoteDetail = activeSheet.Range[activeSheet.Cells[sectionHeaderRow, (int)Columns.BestBidNuked], activeSheet.Cells[sectionHeaderRow, (int)Columns.BestAskRef]];
            Range mergeSophisInputs = activeSheet.Range[activeSheet.Cells[sectionHeaderRow, (int)Columns.UnderlyingPrice], activeSheet.Cells[sectionHeaderRow, (int)Columns.Sicovam]];
            

            //activeSheet.UsedRange.UnMerge();
            mergeMarkComparison.Merge();
            mergeQuoteDetail.Merge();
            mergeSophisInputs.Merge();
            mergeCurrent.Merge();
            mergeEOD.Merge();
            mergeMarkComparison.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            mergeQuoteDetail.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            mergeSophisInputs.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            mergeCurrent.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            mergeEOD.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            mergeMarkComparison.Value2 = "Mark comparison";
            mergeQuoteDetail.Value2 = "Quote Detail";
            mergeSophisInputs.Value2 = "Sophis Inputs";
            mergeCurrent.Value2 = "Today";
            mergeEOD.Value2 = "T-1 EOD";
            mergeMarkComparison.Font.Bold = true;
            mergeQuoteDetail.Font.Bold = true;
            mergeSophisInputs.Font.Bold = true;
            mergeCurrent.Font.Bold = true;
            mergeEOD.Font.Bold = true;
            
            Utils.ReplaceDatatable(activeSheet, convertibleBonds.Count, "Positions", rowCursor, numOfColumns, 1, false);
            foreach (Columns column in Enum.GetValues(typeof(Columns)))
            {
                activeSheet.Cells[rowCursor, column] = columnNames[(int)column];
                activeSheet.Cells[rowCursor, column].Font.Bold = true;
            }
            xlRange = activeSheet.Range[activeSheet.Cells[rowCursor, 1], activeSheet.Cells[rowCursor + convertibleBonds.Count -1, numOfColumns]];

            convertibleBonds.ForEach(cbond =>
            {
                rowCursor++;
                object[,] payloadToApply = new object[1, numOfColumns];
                var cleanIsin = cbond.Isin.Replace("\"", "");
                var sophisPriceSource = markedToLast.Contains(cbond.Isin) ? "Last" : "Theoretical";
                payloadToApply[0, (int)Columns.Isin - 1] = cleanIsin;
                payloadToApply[0, (int)Columns.UnderlyingReference - 1] = cbond.UnderlyingReference;
                payloadToApply[0, (int)Columns.Name - 1] = $@"=BDP([@Isin]& "" Corp"", ""SECURITY_NAME"")";
                payloadToApply[0, (int)Columns.Portfolio - 1] = cbond.Portfolio;
                payloadToApply[0, (int)Columns.Country - 1] = $@"=BDP([@Isin]& "" Corp"", ""CNTRY_OF_RISK"" )";
                payloadToApply[0, (int)Columns.Last - 1] = sophisPriceSource == "Last" ? cbond.Last.ToString() : "";
                payloadToApply[0, (int)Columns.Theoretical - 1] = cbond.Theoretical;
                payloadToApply[0, (int)Columns.LastEOD - 1] = sophisPriceSource == "Last" ? cbond.HistoricalLast.ToString() : "";
                payloadToApply[0, (int)Columns.TheoreticalEOD - 1] = cbond.HistoricalTheoretical;
                payloadToApply[0, (int)Columns.SophisPrice - 1] = $@"=IF([@{columnNames[(int)Columns.SophisPriceSource]}] = ""Last"", [@{columnNames[(int)Columns.Last]}], [@{columnNames[(int)Columns.Theoretical]}] )";
                payloadToApply[0, (int)Columns.SophisPriceEOD - 1] = $@"=IF([@{columnNames[(int)Columns.SophisPriceSource]}] = ""Last"", [@{columnNames[(int)Columns.LastEOD]}], [@{columnNames[(int)Columns.TheoreticalEOD]}] )";
                payloadToApply[0, (int)Columns.SophisPriceSource - 1] = sophisPriceSource;
                payloadToApply[0, (int)Columns.BVAL - 1] = $@"=LET(bdpQuery, BDP([@{columnNames[(int)Columns.Isin]}]& ""@BVAL Corp"", ""LAST_PRICE""), IF(ISNUMBER(bdpQuery), bdpQuery, """"))";
                payloadToApply[0, (int)Columns.BVALEOD - 1] = $@"=LET(bdhQuery, BDH([@Isin] & ""@BVAL Corp"", ""PX_LAST"",  TEXT(IF(WEEKDAY(TODAY())=2, TODAY()-3, TODAY()-1), ""YYYYMMDD"")), IF(ISNUMBER(bdhQuery), bdhQuery, """"))";
                payloadToApply[0, (int)Columns.NukedMid - 1] = $@"=IF(AND(ISNUMBER([@{columnNames[(int)Columns.BestBidNuked]}]), ISNUMBER([@{columnNames[(int)Columns.BestAskNuked]}])), AVERAGE([@{columnNames[(int)Columns.BestBidNuked]}], [@{columnNames[(int)Columns.BestAskNuked]}]), """")";

                payloadToApply[0, (int)Columns.BVALDoDDiff - 1] = $@"=[@{columnNames[(int)Columns.BvalEODDiff]}] - [@{columnNames[(int)Columns.BvalDiff]}]";
                payloadToApply[0, (int)Columns.BvalDiff - 1] = $@"=LET(bvalDiff, [@{columnNames[(int)Columns.SophisPrice]}] - [@{columnNames[(int)Columns.BVAL]}], IF(ISNUMBER(bvalDiff), bvalDiff, """"))";
                payloadToApply[0, (int)Columns.BvalEODDiff - 1] = $@"=LET(bvalEODDiff, [@{columnNames[(int)Columns.SophisPriceEOD]}] - [@{columnNames[(int)Columns.BVALEOD]}] , IF(ISNUMBER(bvalEODDiff), bvalEODDiff, """"))";
                payloadToApply[0, (int)Columns.NukedDiff - 1] = $@"=IF(AND(ISNUMBER([@{columnNames[(int)Columns.NukedMid]}]), ISNUMBER([@{columnNames[(int)Columns.SophisPrice]}])), [@{columnNames[(int)Columns.NukedMid]}] - [@{columnNames[(int)Columns.SophisPrice]}], """")";

                payloadToApply[0, (int)Columns.TotalPNL - 1] = $@"=LET(total, [@{columnNames[(int)Columns.BondPNL]}] + [@{columnNames[(int)Columns.AscotPNL]}], IF(ISNUMBER(total), total, """"))";
                payloadToApply[0, (int)Columns.BondPNL - 1] = $@"=[@{columnNames[(int)Columns.BondNotional]}] * ([@{columnNames[(int)Columns.BvalDiff]}]/100) * -1";
                payloadToApply[0, (int)Columns.AscotPNL - 1] = $@"=MAX([@{columnNames[(int)Columns.BVALMV]}]-[@{columnNames[(int)Columns.SophisMV]}], -[@{columnNames[(int)Columns.SophisMV]}])";

                payloadToApply[0, (int)Columns.TotalEODPNL - 1] = $@"=[@{columnNames[(int)Columns.BondEODPNL]}] + [@{columnNames[(int)Columns.AscotEODPNL]}]";
                payloadToApply[0, (int)Columns.BondEODPNL - 1] = $@"=[@{columnNames[(int)Columns.BondNotionalEOD]}] * ([@{columnNames[(int)Columns.BvalEODDiff]}]/100) * -1";
                payloadToApply[0, (int)Columns.AscotEODPNL - 1] = $@"=MAX([@{columnNames[(int)Columns.BVALEODMV]}]-[@{columnNames[(int)Columns.SophisEODMV]}], -[@{columnNames[(int)Columns.SophisEODMV]}])";

                payloadToApply[0, (int)Columns.BVALMV - 1] = $@"=([@{columnNames[(int)Columns.BVAL]}] - [@{columnNames[(int)Columns.SophisFloor]}])/100 * [@{columnNames[(int)Columns.AscotNotional]}]";
                payloadToApply[0, (int)Columns.BVALEODMV - 1] = $@"=([@{columnNames[(int)Columns.BVALEOD]}] - [@{columnNames[(int)Columns.SophisFloor]}])/100 * [@{columnNames[(int)Columns.AscotNotionalEOD]}]";
                payloadToApply[0, (int)Columns.SophisMV - 1] = $@"=LET(sophisMV, ([@{columnNames[(int)Columns.SophisPrice]}] - [@{columnNames[(int)Columns.SophisFloor]}])/100 * [@{columnNames[(int)Columns.AscotNotional]}], IF(sophisMV<0, 0, sophisMV))";
                payloadToApply[0, (int)Columns.SophisEODMV - 1] = $@"=LET(sophisMV, ([@{columnNames[(int)Columns.SophisPriceEOD]}] - [@{columnNames[(int)Columns.SophisFloor]}])/100 * [@{columnNames[(int)Columns.AscotNotionalEOD]}], IF(sophisMV<0, 0, sophisMV))";
                payloadToApply[0, (int)Columns.BondPosition - 1] = cbond.BondPosition;
                payloadToApply[0, (int)Columns.AscotPosition - 1] = cbond.AscotPosition;
                payloadToApply[0, (int)Columns.BondCurrency - 1] = cbond.Currency;
                payloadToApply[0, (int)Columns.UnderlyingCurrency - 1] = $@"=BDP([@{columnNames[(int)Columns.Isin]}]& ""@BVAL Corp"", ""CV_STOCK_CRNCY"")";
                payloadToApply[0, (int)Columns.FX - 1] = $@"=IF([@{columnNames[(int)Columns.BondCurrency]}] = ""USD"", 1, @BDP(""USD{cbond.Currency} Curncy"", ""LAST_PRICE""))"; 
                payloadToApply[0, (int)Columns.FX_EOD - 1] = $@"=IF([@{columnNames[(int)Columns.BondCurrency]}] = ""USD"", 1, @BDH(""USD{cbond.Currency} Curncy"", ""PX_LAST"",  TEXT(IF(WEEKDAY(TODAY())=2, TODAY()-3, TODAY()-1), ""YYYYMMDD""),  TEXT(IF(WEEKDAY(TODAY())=2, TODAY()-3, TODAY()-1), ""YYYYMMDD"")))"; 
                payloadToApply[0, (int)Columns.BondNotional - 1] = $@"=[@{columnNames[(int)Columns.BondPosition]}]"; // we started using PositionUSD so no need to adjust by FX rate anymore
                payloadToApply[0, (int)Columns.AscotNotional - 1] = $@"=[@{columnNames[(int)Columns.AscotPosition]}]";
                payloadToApply[0, (int)Columns.BondNotionalEOD - 1] = $@"=[@{columnNames[(int)Columns.BondPosition]}]";
                payloadToApply[0, (int)Columns.AscotNotionalEOD - 1] = $@"=[@{columnNames[(int)Columns.AscotPosition]}]";

                var bondFloorBDPQuery = $@"BDP([@{columnNames[(int)Columns.Isin]}]& "" ISIN"",""CV_MODEL_FIXED_INC_VAL"",""ACTIVATE_MULTICALC_MODE"",""Y"",""CV_MODEL_TYP"", ""R"", ""GREEKS_CALCULATION_TYPE"", ""1"", ""CV_MODEL_BORROW_COST"", [@{columnNames[(int)Columns.Borrow]}], ""FLAT_CREDIT_SPREAD_CV_MODEL"", [@{columnNames[(int)Columns.Spread]}], ""CV_MODEL_STOCK_VOL"", [@{columnNames[(int)Columns.Volatility]}], ""CV_MODEL_UNDL_PX"", [@{columnNames[(int)Columns.UnderlyingPrice]}])"; ;
                payloadToApply[0, (int)Columns.SophisFloor - 1] = $@"=LET(BDPQuery, {bondFloorBDPQuery}, IF(ISNUMBER(BDPQuery), BDPQuery, 0))";

                payloadToApply[0, (int)Columns.Par - 1] = $@"=LET(bdpQuery, BDP([@{columnNames[(int)Columns.Isin]}]& ""@BVAL Corp"", ""PAR_AMT""), IF(ISNUMBER(bdpQuery), bdpQuery, """"))"; ;
                payloadToApply[0, (int)Columns.ConvRatio - 1] = $@"=LET(bdpQuery, BDP([@{columnNames[(int)Columns.Isin]}]& ""@BVAL Corp"", ""CV_CNVS_RATIO""), IF(ISNUMBER(bdpQuery), bdpQuery, """"))"; ;
                payloadToApply[0, (int)Columns.Parity - 1] = $@"=(([@{columnNames[(int)Columns.UnderlyingPriceInBondCurrency]}] * [@{columnNames[(int)Columns.ConvRatio]}])/[@{columnNames[(int)Columns.Par]}])*100";
                payloadToApply[0, (int)Columns.Premium - 1] = $@"=([@{columnNames[(int)Columns.SophisPrice]}] - [@{columnNames[(int)Columns.Parity]}])/[@{columnNames[(int)Columns.Parity]}]"; ;
                payloadToApply[0, (int)Columns.Volatility - 1] = cbond.Volatility;
                payloadToApply[0, (int)Columns.ImpliedVol - 1] = $@"=IF(OR(ISBLANK([@{columnNames[(int)Columns.Isin]}]), ISBLANK([@{columnNames[(int)Columns.Borrow]}]), ISBLANK([@{columnNames[(int)Columns.Spread]}]), ISBLANK([@{columnNames[(int)Columns.BestBidRef]}]), ISBLANK([@{columnNames[(int)Columns.BestBid]}]), NOT(ISNUMBER(@BDP([@{columnNames[(int)Columns.Isin]}]&"" ISIN"",""IMPLIED_VOLATILITY_CV"",""ACTIVATE_MULTICALC_MODE"",""Y"",""CV_MODEL_TYP"", ""R"", ""CV_MODEL_BORROW_COST"", [@{columnNames[(int)Columns.Borrow]}], ""FLAT_CREDIT_SPREAD_CV_MODEL"", [@Spread], ""CV_MODEL_UNDL_PX"", [@Best Bid Ref], ""FLAT_FX_VOLATILITY_CV_MODEL"", 0, ""CV_MODEL_BOND_VAL"", [@Best Bid])))), """", @BDP([@Isin]&"" ISIN"",""IMPLIED_VOLATILITY_CV"",""ACTIVATE_MULTICALC_MODE"",""Y"",""CV_MODEL_TYP"", ""R"", ""CV_MODEL_BORROW_COST"", [@Borrow], ""FLAT_CREDIT_SPREAD_CV_MODEL"", [@{columnNames[(int)Columns.Spread]}], ""CV_MODEL_UNDL_PX"", [@{columnNames[(int)Columns.BestBidRef]}], ""FLAT_FX_VOLATILITY_CV_MODEL"", 0, ""CV_MODEL_BOND_VAL"", [@{columnNames[(int)Columns.BestBid]}]))";
                payloadToApply[0, (int)Columns.BestBidNuked - 1] = cbond.BestBidNuked;
                payloadToApply[0, (int)Columns.BestAskNuked - 1] = cbond.BestAskNuked;
                payloadToApply[0, (int)Columns.BestBidIssuer - 1] = cbond.BestBidIssuer;
                payloadToApply[0, (int)Columns.BestBid - 1] = cbond.BestBidBid;
                payloadToApply[0, (int)Columns.BestBidRef - 1] = cbond.BestBidRef;
                payloadToApply[0, (int)Columns.BestAskIssuer - 1] = cbond.BestAskIssuer;
                payloadToApply[0, (int)Columns.BestAsk - 1] = cbond.BestAskAsk;
                payloadToApply[0, (int)Columns.BestAskRef - 1] = cbond.BestAskRef;
                payloadToApply[0, (int)Columns.UnderlyingPrice - 1] = cbond.UnderlyingPrice;
                payloadToApply[0, (int)Columns.UnderlyingPriceInBondCurrency - 1] = $@"=[@{columnNames[(int)Columns.UnderlyingPrice]}] * IF([@{columnNames[(int)Columns.BondCurrency]}] = [@{columnNames[(int)Columns.UnderlyingCurrency]}], 1, BDP([@{columnNames[(int)Columns.UnderlyingCurrency]}]&[@{columnNames[(int)Columns.BondCurrency]}]& "" Curncy"", ""PX_LAST""))";
                payloadToApply[0, (int)Columns.Delta - 1] = cbond.Delta;
                payloadToApply[0, (int)Columns.Spread - 1] = cbond.Spread;
                payloadToApply[0, (int)Columns.Borrow - 1] = cbond.Borrow;
                payloadToApply[0, (int)Columns.Sicovam - 1] = cbond.Sicovam;
                activeSheet.Range[$"A{rowCursor}", $"{Utils.GetExcelColumnName(numOfColumns)}{rowCursor}"].Formula = payloadToApply;
            });

            // Add cell with total for following rows
            Array.ForEach(new[]
            {
                (int)Columns.BondPNL,
                (int)Columns.AscotPNL,
                (int)Columns.TotalPNL,
                (int)Columns.BondEODPNL,
                (int)Columns.AscotEODPNL,
                (int)Columns.TotalEODPNL,
            }, column =>
            {
                string startCell = (activeSheet.Cells[rowCursor - convertibleBonds.Count + 1, column]).Address[false, false];
                string endCell = (activeSheet.Cells[rowCursor, column]).Address[false, false];
                var totalCell = activeSheet.Cells[rowCursor + 1, column];
                totalCell.Formula = $"=SUBTOTAL(9,{startCell}:{endCell})";
            });

            Array.ForEach(new[]
            {
                (int)Columns.BVAL,
                (int)Columns.BVALEOD,
                (int)Columns.BestBid,
                (int)Columns.BestAsk,
                (int)Columns.SophisPrice,
                (int)Columns.SophisPriceEOD,
                (int)Columns.UnderlyingPrice,
                (int)Columns.UnderlyingPriceInBondCurrency,
                (int)Columns.ImpliedVol,
                (int)Columns.BvalDiff,
                (int)Columns.BvalEODDiff,
                (int)Columns.NukedDiff,
                (int)Columns.NukedMid,
                (int)Columns.BestAskNuked,
                (int)Columns.BestBidNuked,
                (int)Columns.Delta,
                (int)Columns.FX,
                (int)Columns.FX_EOD,
                (int)Columns.BVALDoDDiff,
                (int)Columns.Parity,
                (int)Columns.SophisFloor,
                (int)Columns.ConvRatio,
            }, column => activeSheet.Columns[column].NumberFormat = "0.00");

            ApplyBuiltInConditionalFormatting(rowCursor);

            activeSheet.Columns[(int)Columns.Premium].NumberFormat = "#,##0.00%_);[Red](#,##0.00%)";
            activeSheet.Columns[(int)Columns.SophisMV].NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            activeSheet.Columns[(int)Columns.SophisEODMV].NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            activeSheet.Columns[(int)Columns.BVALMV].NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            activeSheet.Columns[(int)Columns.BVALEODMV].NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            activeSheet.Columns[(int)Columns.BondPosition].NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            activeSheet.Columns[(int)Columns.AscotPosition].NumberFormat = "#,##0.00_);[Red](#,##0.00)";
            activeSheet.Columns[(int)Columns.BondNotional].NumberFormat = "[Blue]$#,##0_);[Red]$(#,##0)";
            activeSheet.Columns[(int)Columns.BondNotionalEOD].NumberFormat = "[Blue]$#,##0_);[Red]$(#,##0)";
            activeSheet.Columns[(int)Columns.AscotNotional].NumberFormat = "[Blue]$#,##0_);[Red]$(#,##0)";
            activeSheet.Columns[(int)Columns.AscotNotionalEOD].NumberFormat = "[Blue]$#,##0_);[Red]$(#,##0)";
            activeSheet.Columns[(int)Columns.BondPNL].NumberFormat = "[Blue]$#,##0_);[Red]$(#,##0)";
            activeSheet.Columns[(int)Columns.BondEODPNL].NumberFormat = "[Blue]$#,##0_);[Red]$(#,##0)";
            activeSheet.Columns[(int)Columns.AscotPNL].NumberFormat = "[Blue]$#,##0_);[Red]$(#,##0)";
            activeSheet.Columns[(int)Columns.AscotEODPNL].NumberFormat = "[Blue]$#,##0_);[Red]$(#,##0)";
            activeSheet.Columns[(int)Columns.TotalPNL].NumberFormat = "[Blue]$#,##0_);[Red]$(#,##0)";
            activeSheet.Columns[(int)Columns.TotalEODPNL].NumberFormat = "[Blue]$#,##0_);[Red]$(#,##0)";
            activeSheet.Columns[(int)Columns.Par].NumberFormat = "#,##0";

            activeSheet.Range[activeSheet.Columns[1], activeSheet.Columns[numOfColumns]].AutoFit();
            activeSheet.Columns[Columns.Isin].ColumnWidth = 14.5;
            activeSheet.Columns[Columns.Par].ColumnWidth = 9.5;
            activeSheet.Columns[Columns.Country].ColumnWidth = 9.5;
            activeSheet.Columns[Columns.FX].ColumnWidth = 9.5;
            activeSheet.Columns[Columns.FX_EOD].ColumnWidth = 9.5;
            activeSheet.Columns[Columns.BondCurrency].ColumnWidth = 9.5;
            activeSheet.Columns[Columns.UnderlyingCurrency].ColumnWidth = 9.5;
            activeSheet.Columns[Columns.BondPNL].ColumnWidth = 15;
            activeSheet.Columns[Columns.BondEODPNL].ColumnWidth = 15;
            activeSheet.Columns[Columns.AscotPNL].ColumnWidth = 15;
            activeSheet.Columns[Columns.AscotEODPNL].ColumnWidth = 15;
            activeSheet.Columns[Columns.TotalPNL].ColumnWidth = 16;
            activeSheet.Columns[Columns.TotalEODPNL].ColumnWidth = 16;

            // Remove all borders
            activeSheet.UsedRange.Borders.LineStyle = XlLineStyle.xlLineStyleNone;
            
            // Add borders to EOD and Current ranges
            var eodRange = activeSheet.Range[activeSheet.Cells[sectionHeaderRow, (int)Columns.SophisPriceEOD], activeSheet.Cells[rowCursor + 1, (int)Columns.AscotEODPNL]];
            Borders eodRangeBorder = eodRange.Borders;
            eodRangeBorder[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            eodRangeBorder[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThick;

            eodRangeBorder[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            eodRangeBorder[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThick;
            
            eodRangeBorder[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            eodRangeBorder[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;

            var currentRange = activeSheet.Range[activeSheet.Cells[sectionHeaderRow, (int)Columns.SophisPrice], activeSheet.Cells[rowCursor + 1, (int)Columns.AscotPNL]];
            Borders currentRangeBorder = currentRange.Borders;
            currentRangeBorder[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            currentRangeBorder[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThick;

            currentRangeBorder[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            currentRangeBorder[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThick;

            currentRangeBorder[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            currentRangeBorder[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;
            
            Array.ForEach(new[]
            {
                (int)Columns.Separator1,
                (int)Columns.Separator2,
                (int)Columns.Separator3,
            }, column => {
                activeSheet.Columns[column].Interior.Color = XlRgbColor.rgbWhite;
                activeSheet.Columns[column].ColumnWidth = 3;
                activeSheet.Columns[column].Borders.Color = XlRgbColor.rgbWhite;
            });

            var separator2Range = activeSheet.Range[activeSheet.Cells[sectionHeaderRow, (int)Columns.Separator2], activeSheet.Cells[rowCursor + 1, (int)Columns.Separator2]];
            Borders separator2Border = separator2Range.Borders;
            separator2Border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            separator2Border[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThick;
            separator2Border[XlBordersIndex.xlEdgeLeft].Color = XlRgbColor.rgbBlack;

            var separator3Range = activeSheet.Range[activeSheet.Cells[sectionHeaderRow, (int)Columns.Separator3], activeSheet.Cells[rowCursor + 1, (int)Columns.Separator3]];
            Borders separator3Border = separator3Range.Borders;
            separator3Border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            separator3Border[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThick;
            separator3Border[XlBordersIndex.xlEdgeLeft].Color = XlRgbColor.rgbBlack;

            // First ungroup all columns
            for (int i = 1; i <= numOfColumns; i++)
            {
                Range column = activeSheet.Columns[i];
                while (column.OutlineLevel > 1)
                {
                    try
                    {
                        column.EntireColumn.Ungroup();
                    }
                    catch (COMException)
                    {
                        break;
                    }
                }
            }
            //Then group required columns
            var columnsToGroup = new List<int>
                    {
                        (int)Columns.Isin,
                        (int)Columns.FX,
                        (int)Columns.FX_EOD,
                        (int)Columns.BondCurrency,
                        (int)Columns.UnderlyingCurrency,
                        (int)Columns.UnderlyingPriceInBondCurrency,
                        (int)Columns.Theoretical,
                        (int)Columns.Last,
                        (int)Columns.TheoreticalEOD,
                        (int)Columns.LastEOD,
                        (int)Columns.BVALMV,
                        (int)Columns.BVALEODMV,
                        (int)Columns.SophisMV,
                        (int)Columns.SophisEODMV,
                        (int)Columns.BondPosition,
                        (int)Columns.BondNotional,
                        (int)Columns.BondNotionalEOD,
                        (int)Columns.AscotPosition,
                        (int)Columns.AscotNotional,
                        (int)Columns.AscotNotionalEOD,
                        (int)Columns.BondEODPNL,
                        (int)Columns.AscotEODPNL,
                        (int)Columns.BondPNL,
                        (int)Columns.AscotPNL,
                        (int)Columns.ConvRatio,
                        (int)Columns.Par,
                    };
            columnsToGroup.ForEach(col =>
            {
                var range = activeSheet.Columns[col];
                range.Group();
                range.EntireColumn.Hidden = true;
            });

            // Freeze first 3 columns
            Range cell = activeSheet.Cells[1, 3 + 1];
            cell.Select();
            activeSheet.Application.ActiveWindow.FreezePanes = true;
        }
    }
}
