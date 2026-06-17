using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SWTableConvertor
{
    public static class SolidWorksInserter
    {
        public static string InsertIntoActiveDrawing(
            ExcelTable table, 
            bool firstRowIsHeader, 
            bool importFontStyles, 
            bool importAlignment, 
            Action<string> log)
        {
            if (table == null || table.Headers.Count == 0)
                throw new InvalidOperationException("No table data to insert.");

            log("Connecting to SOLIDWORKS...");
            ISldWorks sw;
            try
            {
                sw = (ISldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch (COMException)
            {
                throw new InvalidOperationException("SOLIDWORKS is not running. Please start SOLIDWORKS and open a drawing document.");
            }

            IModelDoc2 model = sw.ActiveDoc as IModelDoc2;
            if (model == null)
                throw new InvalidOperationException("No active document found in SOLIDWORKS.");

            if (model.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                throw new InvalidOperationException("The active SOLIDWORKS document is not a drawing (.slddrw). Please open or activate a drawing.");

            IDrawingDoc drawing = model as IDrawingDoc;
            if (drawing == null)
                throw new InvalidOperationException("Active document cannot be accessed as a drawing.");

            string title = model.GetTitle();
            string path = model.GetPathName();
            bool wasDirty = model.GetSaveFlag();

            ISheet sheet = drawing.GetCurrentSheet() as ISheet;
            string sheetName = sheet != null ? sheet.GetName() : "";

            int rowCount = table.Rows.Count + (firstRowIsHeader ? 1 : 0);
            int colCount = table.Headers.Count;

            log(string.Format("Inserting general table with {0} rows and {1} columns...", rowCount, colCount));

            IModelDocExtension ext = model.Extension;
            
            // Insert at 0.02, 0.26 (typical placement on A4/A3 sheets)
            TableAnnotation tableAnnotation = ext.InsertGeneralTableAnnotation(
                false, // UseAnchorPoint (false = place at coordinates)
                0.02,  // X coordinate (meters)
                0.26,  // Y coordinate (meters)
                (int)swTableHeaderPosition_e.swTableHeader_Top,
                "",    // TemplatePath (empty = default general table)
                rowCount,
                colCount);

            if (tableAnnotation == null)
                throw new InvalidOperationException("SOLIDWORKS failed to create the general table. Ensure you are in sheet edit mode and drawing is not read-only.");

            log("Populating table cells and transferring formats...");
            FillTable(tableAnnotation, table, firstRowIsHeader, importFontStyles, importAlignment, log);

            string dirtyText = wasDirty ? "drawing was already modified" : "drawing is now modified";
            return string.Format("Successfully inserted table into active drawing '{0}'{1}. {2}.", 
                title, 
                string.IsNullOrEmpty(sheetName) ? "" : " on sheet '" + sheetName + "'", 
                dirtyText);
        }

        private static void FillTable(
            TableAnnotation tableAnnotation, 
            ExcelTable table, 
            bool firstRowIsHeader,
            bool importFontStyles,
            bool importAlignment,
            Action<string> log)
        {
            int colCount = table.Headers.Count;
            int totalRows = table.Rows.Count + (firstRowIsHeader ? 1 : 0);

            // 1. Write Headers (Row 0)
            for (int c = 0; c < colCount; c++)
            {
                ExcelCell cell = table.Headers[c];
                tableAnnotation.set_Text2(0, c, false, cell.Value);
                ApplyFormatting(tableAnnotation, 0, c, cell, importFontStyles, importAlignment);
            }

            // 2. Write Data Rows (Row 1 to N)
            for (int r = 0; r < table.Rows.Count; r++)
            {
                int targetRow = r + (firstRowIsHeader ? 1 : 0);
                for (int c = 0; c < colCount; c++)
                {
                    ExcelCell cell = table.Rows[r][c];
                    tableAnnotation.set_Text2(targetRow, c, false, cell.Value);
                    ApplyFormatting(tableAnnotation, targetRow, c, cell, importFontStyles, importAlignment);
                }
            }

            // 3. Set fallback alignment if alignment importing is off
            if (!importAlignment)
            {
                try
                {
                    tableAnnotation.TextHorizontalJustification = (int)swTextJustification_e.swTextJustificationCenter;
                    tableAnnotation.TextVerticalJustification = (int)swVerticalJustification_e.swVerticalJustificationMiddle;
                }
                catch { }
            }

            // 4. Calculate and set smart column widths
            double[] colWidths = new double[colCount];
            for (int c = 0; c < colCount; c++)
            {
                int maxChars = table.Headers[c].Value.Length;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    int len = table.Rows[r][c].Value.Length;
                    if (len > maxChars) maxChars = len;
                }

                // Map char count to meters: Min 18mm, Max 150mm. Each char is ~2.2mm + 6mm padding.
                double width = Math.Max(0.018, Math.Min(0.150, (maxChars * 0.0022) + 0.006));
                colWidths[c] = width;
            }

            for (int c = 0; c < colCount; c++)
            {
                try
                {
                    tableAnnotation.SetColumnWidth(c, colWidths[c], (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
                }
                catch { }
            }

            // 5. Set default row heights (7.5mm / 0.0075m)
            for (int r = 0; r < totalRows; r++)
            {
                try
                {
                    tableAnnotation.SetRowHeight(r, 0.0075, (int)swTableRowColSizeChangeBehavior_e.swTableRowColChange_TableSizeCanChange);
                }
                catch { }
            }
        }

        private static void ApplyFormatting(
            TableAnnotation tableAnnotation, 
            int row, 
            int col, 
            ExcelCell cell,
            bool importFontStyles,
            bool importAlignment)
        {
            // Set cell alignment
            if (importAlignment)
            {
                try
                {
                    int justification = (int)swTextJustification_e.swTextJustificationCenter; // default
                    if (cell.Alignment == "left")
                    {
                        justification = (int)swTextJustification_e.swTextJustificationLeft;
                    }
                    else if (cell.Alignment == "right")
                    {
                        justification = (int)swTextJustification_e.swTextJustificationRight;
                    }

                    tableAnnotation.set_CellTextHorizontalJustification(row, col, justification);
                    tableAnnotation.set_CellTextVerticalJustification(row, col, (int)swVerticalJustification_e.swVerticalJustificationMiddle);
                }
                catch { }
            }

            // Set cell font (Bold/Italic)
            if (importFontStyles)
            {
                if (cell.Bold || cell.Italic)
                {
                    try
                    {
                        dynamic textFormat = tableAnnotation.GetCellTextFormat(row, col);
                        if (textFormat != null)
                        {
                            textFormat.Bold = cell.Bold;
                            textFormat.Italic = cell.Italic;
                            tableAnnotation.SetCellTextFormat(row, col, false, textFormat);
                        }
                    }
                    catch { }
                }
            }
        }
    }
}
