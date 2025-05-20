using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using MarkingSheet.Sheets;
using static System.String;
using static MarkingSheet.Mark.MarkCdsUtil;
using System.Diagnostics;

namespace MarkingSheet
{
    // https://www.flaticon.com/free-icon-font
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        

        private void loadPositionsButton_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            activeSheet.Activate();

            if (activeSheet.Cells[1, 1].Value == null || activeSheet.Cells[1, 1].Value?.ToString() == "" || activeSheet.Cells[1, 1].Value?.ToString() == "Bond Marks")
            {
                BondMarkingSheet.LoadBondMarkingSheetPositions();
            }
            else
            {
                MessageBox.Show("Please open Bonds marks tab to refresh Bond Positions.");
                return;
            }
        }

        private void sendMarksToSophisButton_Click(object sender, RibbonControlEventArgs e)
        {
            var activeSheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);
            if (activeSheet.Cells[1, 1].Value?.ToString() != "Bond Marks")
            {
                MessageBox.Show("Please open Bond marks tab to submit Bond marks");
                return;
            }

            var availableTables = Utils.GetDataTableContentsRaw(activeSheet);

            var positionsTable = availableTables.Where(table => table.Name == "Positions").ToList();

            if (positionsTable != null && positionsTable.Count() == 0) 
            {
                MessageBox.Show("No instruments available for marking");
                return;
            }

            var invalidMarks = new List<String>();

            var readMarks = positionsTable.First().Rows.Select(row =>
            {
                if (!int.TryParse(row["Sicovam"].ToString(), out int sicovam))
                {
                    invalidMarks.Add($"Invalid Mark to Upload for ${row["Sicovam"]} - ${row["Mark to Upload"]}");
                }

                if (!double.TryParse(row["Mark to Upload"].ToString(), out double level))
                {
                    invalidMarks.Add($"Invalid Mark to Upload for ${row["Sicovam"]} - ${row["Mark to Upload"]}");
                }
            
                return new MarkBondUtil.Mark()
                {
                    sicovam = sicovam,
                    level = level,
                    fieldName = "Last"
                };
            });

            var theoMarks = positionsTable.First().Rows.Select(row =>
            {
                if (!int.TryParse(row["Sicovam"].ToString(), out int sicovam))
                {
                    invalidMarks.Add($"Invalid Mark to Upload for ${row["Sicovam"]} - ${row["Mark to Upload"]}");
                }

                if (!double.TryParse(row["Mark to Upload"].ToString(), out double level))
                {
                    invalidMarks.Add($"Invalid Mark to Upload for ${row["Sicovam"]} - ${row["Mark to Upload"]}");
                }
                return new MarkBondUtil.Mark()
                {
                    sicovam = sicovam,
                    level = level,
                    fieldName = "T"
                };
            });

            if (invalidMarks.Count > 0)
            {
                MessageBox.Show(Join("\n", invalidMarks));
                return;
            }

            MarkBondUtil.SubmitMarks(readMarks.Concat(theoMarks));
        }

        private void loadCdsPositions_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
                activeSheet.Activate();

                if (activeSheet.Cells[1, 1].Value == null || activeSheet.Cells[1, 1].Value?.ToString() == "" ||
                    activeSheet.Cells[1, 1].Value?.ToString() == "CDS Marks")
                {
                    CdsMarkingSheet.LoadCdsMarkingSheetPositions();
                }
                else
                {
                    MessageBox.Show("Please open CDS marks tab to refresh CDS Positions.");
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void sendCdsMarksButton_Click(object sender, RibbonControlEventArgs e)
        {

            var activeSheet = ((Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet);


            if (activeSheet.Cells[1, 1].Value?.ToString() != "CDS Marks")
            {
                MessageBox.Show("Please open CDS marks tab to submit CDS marks");
                return;
            }


            var curves = CdsMarkingSheet.GetCurves(activeSheet, "CDSCurves");

            var byReference = curves.GroupBy(x => x.Ticker);

            var cdsMarkResult = byReference.Select(curveGroup =>
            {
                var points = new List<CdsPoint>();

                // Each curve needs to have all points from all seniorities submitted at once
                foreach (var curve in curveGroup)
                {
                    if (curve.OneYear.HasValue)
                    {
                        points.Add(new CdsPoint()
                        {
                            PeriodMultiplier = 1.0,
                            Rate = curve.OneYear.Value,
                            Seniority = curve.Seniority
                        });
                    }

                    if (curve.ThreeYear.HasValue)
                    {
                        points.Add(new CdsPoint()
                        {
                            PeriodMultiplier = 3.0,
                            Rate = curve.ThreeYear.Value,
                            Seniority = curve.Seniority
                        });
                    }

                    if (curve.FiveYear.HasValue)
                    {
                        points.Add(new CdsPoint()
                        {
                            PeriodMultiplier = 5.0,
                            Rate = curve.FiveYear.Value,
                            Seniority = curve.Seniority
                        });
                    }

                    if (curve.SevenYear.HasValue)
                    {
                        points.Add(new CdsPoint()
                        {
                            PeriodMultiplier = 7.0,
                            Rate = curve.SevenYear.Value,
                            Seniority = curve.Seniority
                        });
                    }

                    if (curve.TenYear.HasValue)
                    {
                        points.Add(new CdsPoint()
                        {
                            PeriodMultiplier = 10.0,
                            Rate = curve.TenYear.Value,
                            Seniority = curve.Seniority
                        });
                    }
                }


                if (points.Count > 0)
                {
                    return SubmitMarks(new CdsMarkParameters()
                    {
                        Reference = curveGroup.Key,
                        Points = points
                    });
                }

                return "Skipped";
            });

            var errorMessages = cdsMarkResult.Where(x => x != "OK" && x != "Skipped").ToList();
            MessageBox.Show(errorMessages.Any() ? Join("\n", errorMessages) : "CDS marks submitted successfully");
        }

        private void loadCbPositionsButton_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;
            activeSheet.Activate();

            try
            {
                if (activeSheet.Cells[1, 1].Value == null || activeSheet.Cells[1, 1].Value?.ToString() == "" ||
                    activeSheet.Cells[1, 1].Value?.ToString() == "Convertible Bond Marks")
                {
                    activeSheet.Cells[1, 1].Value = "Loading...";
                    CBondMarkingSheet.LoadCBondMarkingSheetPositions();
                }
                else
                {
                    MessageBox.Show("Please open Convertible Bonds marks tab to refresh CB Positions.");
                    return;
                }
            }
            catch (Exception ex) // Catching general exception, consider catching more specific exceptions if possible
            {
                var errorMessage = $"An error occurred: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}";
                MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void bbgButton_Click(object sender, RibbonControlEventArgs e)
        {
            var button = (RibbonButton)sender;
            TriggerBloomberg(button.Label);
        }

        private void TriggerBloomberg(string item)
        {
            if (item == "")
            {
                return;
            }

            var worksheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet;

            var selectedRow = (Range)Globals.ThisAddIn.Application.Selection;

            var row = selectedRow.Cells.Row;

            string[] equityFunctions = { "HVG", "G3", "G7", "G8", "G10", "DVD" };

            var reference = worksheet.Cells[row, CBondMarkingSheet.Columns.Isin].Value;
            string search = $"{reference} CORP";

            if (equityFunctions.Contains(item)){
                search = worksheet.Cells[row, CBondMarkingSheet.Columns.UnderlyingReference].Value;
            }

            Process.Start(new ProcessStartInfo($"bbg://securities/{search}/{item}")
            { UseShellExecute = true });
        }
    }
}
