using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml;

namespace SWTableConvertor
{
    public class ExcelCell
    {
        public string Value { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public string Alignment { get; set; } // "left", "center", "right", "general"

        public ExcelCell()
        {
            Value = "";
            Alignment = "general";
        }

        public ExcelCell(string val)
        {
            Value = val ?? "";
            Alignment = "general";
        }
    }

    public class ExcelSheetInfo
    {
        public string Name { get; set; }
        public string TargetPath { get; set; }
    }

    public class ExcelTable
    {
        public string RangeAddress { get; set; }
        public List<ExcelCell> Headers { get; set; }
        public List<List<ExcelCell>> Rows { get; set; }

        public ExcelTable()
        {
            Headers = new List<ExcelCell>();
            Rows = new List<List<ExcelCell>>();
        }
    }

    public static class ExcelReader
    {
        private class ExcelFont
        {
            public bool Bold { get; set; }
            public bool Italic { get; set; }
        }

        private class ExcelStyle
        {
            public bool Bold { get; set; }
            public bool Italic { get; set; }
            public string Alignment { get; set; }
        }

        private static string GetAttrValue(XmlNode node, string name)
        {
            if (node == null || node.Attributes == null) return null;
            XmlAttribute attr = node.Attributes[name];
            return attr != null ? attr.Value : null;
        }

        private static string GetAttrValueNS(XmlNode node, string name, string ns)
        {
            if (node == null || node.Attributes == null) return null;
            XmlAttribute attr = node.Attributes[name, ns];
            return attr != null ? attr.Value : null;
        }

        public static List<ExcelSheetInfo> GetSheets(string xlsxPath)
        {
            var sheets = new List<ExcelSheetInfo>();
            if (!File.Exists(xlsxPath))
                throw new FileNotFoundException("Excel file not found.", xlsxPath);

            using (ZipArchive archive = ZipFile.OpenRead(xlsxPath))
            {
                var rels = new Dictionary<string, string>();
                ZipArchiveEntry relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
                if (relsEntry != null)
                {
                    using (Stream stream = relsEntry.Open())
                    {
                        XmlDocument doc = new XmlDocument();
                        doc.Load(stream);
                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                        nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");

                        XmlNodeList relNodes = doc.SelectNodes("//r:Relationship", nsmgr);
                        if (relNodes != null)
                        {
                            foreach (XmlNode node in relNodes)
                            {
                                string id = GetAttrValue(node, "Id");
                                string target = GetAttrValue(node, "Target");
                                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                                {
                                    if (!target.StartsWith("xl/"))
                                    {
                                        if (target.StartsWith("/"))
                                            target = "xl" + target;
                                        else
                                            target = "xl/" + target;
                                    }
                                    rels[id] = target;
                                }
                            }
                        }
                    }
                }

                ZipArchiveEntry workbookEntry = archive.GetEntry("xl/workbook.xml");
                if (workbookEntry != null)
                {
                    using (Stream stream = workbookEntry.Open())
                    {
                        XmlDocument doc = new XmlDocument();
                        doc.Load(stream);
                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                        nsmgr.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                        nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                        XmlNodeList sheetNodes = doc.SelectNodes("//d:sheet", nsmgr);
                        if (sheetNodes != null)
                        {
                            foreach (XmlNode node in sheetNodes)
                            {
                                string name = GetAttrValue(node, "name");
                                string rId = GetAttrValueNS(node, "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                                if (string.IsNullOrEmpty(rId))
                                {
                                    rId = GetAttrValue(node, "r:id");
                                }

                                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(rId) && rels.ContainsKey(rId))
                                {
                                    sheets.Add(new ExcelSheetInfo
                                    {
                                        Name = name,
                                        TargetPath = rels[rId]
                                    });
                                }
                            }
                        }
                    }
                }
            }

            return sheets;
        }

        private static List<ExcelStyle> ReadStyles(ZipArchive archive)
        {
            var styles = new List<ExcelStyle>();
            ZipArchiveEntry entry = archive.GetEntry("xl/styles.xml");
            if (entry == null) return styles;

            using (Stream stream = entry.Open())
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(stream);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                // 1. Read Fonts
                var fonts = new List<ExcelFont>();
                XmlNodeList fontNodes = doc.SelectNodes("//d:fonts/d:font", nsmgr);
                if (fontNodes != null)
                {
                    foreach (XmlNode fNode in fontNodes)
                    {
                        bool bold = fNode.SelectSingleNode("d:b", nsmgr) != null;
                        bool italic = fNode.SelectSingleNode("d:i", nsmgr) != null;
                        fonts.Add(new ExcelFont { Bold = bold, Italic = italic });
                    }
                }

                // 2. Read cell formatting (cellXfs)
                XmlNodeList xfNodes = doc.SelectNodes("//d:cellXfs/d:xf", nsmgr);
                if (xfNodes != null)
                {
                    foreach (XmlNode xfNode in xfNodes)
                    {
                        var style = new ExcelStyle();

                        string fontIdStr = GetAttrValue(xfNode, "fontId");
                        int fontId;
                        if (!string.IsNullOrEmpty(fontIdStr) && int.TryParse(fontIdStr, out fontId) && fontId >= 0 && fontId < fonts.Count)
                        {
                            style.Bold = fonts[fontId].Bold;
                            style.Italic = fonts[fontId].Italic;
                        }

                        XmlNode alignNode = xfNode.SelectSingleNode("d:alignment", nsmgr);
                        if (alignNode != null)
                        {
                            style.Alignment = GetAttrValue(alignNode, "horizontal") ?? "general";
                        }
                        else
                        {
                            style.Alignment = "general";
                        }

                        styles.Add(style);
                    }
                }
            }

            return styles;
        }

        public static ExcelTable ReadSheet(string xlsxPath, ExcelSheetInfo sheetInfo, string rangeStr = null, bool firstRowIsHeader = true)
        {
            if (!File.Exists(xlsxPath))
                throw new FileNotFoundException("Excel file not found.", xlsxPath);

            using (ZipArchive archive = ZipFile.OpenRead(xlsxPath))
            {
                // 1. Read Shared Strings & Styles
                List<string> sharedStrings = ReadSharedStrings(archive);
                List<ExcelStyle> styles = ReadStyles(archive);

                // 2. Open sheet entry
                ZipArchiveEntry sheetEntry = archive.GetEntry(sheetInfo.TargetPath);
                if (sheetEntry == null)
                    throw new FileNotFoundException("Worksheet data not found in ZIP archive: " + sheetInfo.TargetPath);

                var cells = new Dictionary<string, ExcelCell>();
                int minRow = int.MaxValue, maxRow = int.MinValue;
                int minCol = int.MaxValue, maxCol = int.MinValue;

                using (Stream stream = sheetEntry.Open())
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(stream);
                    XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                    nsmgr.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                    XmlNodeList cellNodes = doc.SelectNodes("//d:c", nsmgr);
                    if (cellNodes != null)
                    {
                        foreach (XmlNode node in cellNodes)
                        {
                            string cellRef = GetAttrValue(node, "r");
                            if (string.IsNullOrEmpty(cellRef)) continue;

                            string type = GetAttrValue(node, "t");
                            string val = "";

                            XmlNode vNode = node.SelectSingleNode("d:v", nsmgr);
                            if (vNode != null)
                            {
                                val = vNode.InnerText;
                            }

                            if (type == "s") // Shared string
                            {
                                int idx;
                                if (int.TryParse(val, out idx) && idx >= 0 && idx < sharedStrings.Count)
                                {
                                    val = sharedStrings[idx];
                                }
                            }
                            else if (type == "inlineStr") // Inline string
                            {
                                XmlNode tNode = node.SelectSingleNode("d:is/d:t", nsmgr);
                                if (tNode != null)
                                {
                                    val = tNode.InnerText;
                                }
                            }
                            else if (type == "str") // Formula calculated string
                            {
                                // val has the text value directly
                            }
                            else if (type == "b") // Boolean
                            {
                                val = (val == "1" || val.ToLower() == "true") ? "TRUE" : "FALSE";
                            }

                            // Read style index
                            string styleIdxStr = GetAttrValue(node, "s");
                            int styleIdx;
                            bool bold = false;
                            bool italic = false;
                            string alignment = "general";

                            if (!string.IsNullOrEmpty(styleIdxStr) && int.TryParse(styleIdxStr, out styleIdx) && styleIdx >= 0 && styleIdx < styles.Count)
                            {
                                bold = styles[styleIdx].Bold;
                                italic = styles[styleIdx].Italic;
                                alignment = styles[styleIdx].Alignment;
                            }

                            var cell = new ExcelCell
                            {
                                Value = val ?? "",
                                Bold = bold,
                                Italic = italic,
                                Alignment = alignment
                            };

                            cells[cellRef] = cell;

                            // Update bounding box if cell has content
                            if (!string.IsNullOrEmpty(val))
                            {
                                int r, c;
                                ParseCellReference(cellRef, out r, out c);
                                if (r < minRow) minRow = r;
                                if (r > maxRow) maxRow = r;
                                if (c < minCol) minCol = c;
                                if (c > maxCol) maxCol = c;
                            }
                        }
                    }
                }

                // Determine range bounds
                int startRow, startCol, endRow, endCol;
                if (!string.IsNullOrEmpty(rangeStr) && TryParseRange(rangeStr, out startRow, out startCol, out endRow, out endCol))
                {
                    // Use user-defined range
                }
                else
                {
                    if (minRow == int.MaxValue) // Empty sheet
                    {
                        return new ExcelTable { RangeAddress = "Empty" };
                    }
                    startRow = minRow;
                    startCol = minCol;
                    endRow = maxRow;
                    endCol = maxCol;
                }

                var table = new ExcelTable();
                table.RangeAddress = GetCellReference(startRow, startCol) + ":" + GetCellReference(endRow, endCol);

                int rowsCount = endRow - startRow + 1;
                int colsCount = endCol - startCol + 1;

                if (rowsCount <= 0 || colsCount <= 0)
                    return table;

                // Load rows
                int currentRowIndex = startRow;
                if (firstRowIsHeader)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        string refStr = GetCellReference(currentRowIndex, col);
                        ExcelCell cellVal;
                        if (!cells.TryGetValue(refStr, out cellVal))
                        {
                            cellVal = new ExcelCell();
                        }
                        table.Headers.Add(cellVal);
                    }
                    currentRowIndex++;
                }
                else
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        table.Headers.Add(new ExcelCell("Column " + GetColumnName(col)));
                    }
                }

                for (; currentRowIndex <= endRow; currentRowIndex++)
                {
                    var rowData = new List<ExcelCell>();
                    for (int col = startCol; col <= endCol; col++)
                    {
                        string refStr = GetCellReference(currentRowIndex, col);
                        ExcelCell cellVal;
                        if (!cells.TryGetValue(refStr, out cellVal))
                        {
                            cellVal = new ExcelCell();
                        }
                        rowData.Add(cellVal);
                    }
                    table.Rows.Add(rowData);
                }

                return table;
            }
        }

        private static List<string> ReadSharedStrings(ZipArchive archive)
        {
            var list = new List<string>();
            ZipArchiveEntry entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null) return list;

            using (Stream stream = entry.Open())
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(stream);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                XmlNodeList siNodes = doc.SelectNodes("//d:si", nsmgr);
                if (siNodes != null)
                {
                    foreach (XmlNode siNode in siNodes)
                    {
                        list.Add(siNode.InnerText ?? "");
                    }
                }
            }
            return list;
        }

        public static void ParseCellReference(string cellRef, out int row, out int col)
        {
            cellRef = cellRef.ToUpperInvariant();
            int colLettersCount = 0;
            while (colLettersCount < cellRef.Length && char.IsLetter(cellRef[colLettersCount]))
            {
                colLettersCount++;
            }
            string colStr = cellRef.Substring(0, colLettersCount);
            string rowStr = cellRef.Substring(colLettersCount);

            row = int.Parse(rowStr, CultureInfo.InvariantCulture) - 1; // 0-based

            col = 0;
            for (int i = 0; i < colStr.Length; i++)
            {
                col = col * 26 + (colStr[i] - 'A' + 1);
            }
            col = col - 1; // 0-based
        }

        public static string GetCellReference(int row, int col)
        {
            return GetColumnName(col) + (row + 1).ToString(CultureInfo.InvariantCulture);
        }

        public static string GetColumnName(int col)
        {
            int dividend = col + 1;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = (char)(65 + modulo) + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static bool TryParseRange(string rangeStr, out int startRow, out int startCol, out int endRow, out int endCol)
        {
            startRow = startCol = endRow = endCol = -1;
            if (string.IsNullOrEmpty(rangeStr)) return false;

            string[] parts = rangeStr.Split(':');
            if (parts.Length != 2) return false;

            try
            {
                ParseCellReference(parts[0].Trim(), out startRow, out startCol);
                ParseCellReference(parts[1].Trim(), out endRow, out endCol);

                if (startRow > endRow) { int t = startRow; startRow = endRow; endRow = t; }
                if (startCol > endCol) { int t = startCol; startCol = endCol; endCol = t; }
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
