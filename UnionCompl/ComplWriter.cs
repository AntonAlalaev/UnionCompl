using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Style;
using System.Drawing;

namespace UnionCompl
{
    internal class ComplWriter
    {
        public ComplWriter()
        {
            
        }

        public void CopyRowsWithFormatting(string sourceFilePath, string targetFilePath, int rowStart, int rowEnd)
        {
            if (rowStart < 1 || rowEnd < 1 || rowEnd < rowStart)
            {
                throw new ArgumentException("Invalid row numbers. Both must be positive and rowEnd must be greater than or equal to rowStart.");
            }

            if (!File.Exists(sourceFilePath))
            {
                throw new FileNotFoundException("Source file not found.", sourceFilePath);
            }

            using (var package = new ExcelPackage(new FileInfo(sourceFilePath)))
            {
                ExcelWorkbook sourceWorkbook = package.Workbook;
                if (sourceWorkbook.Worksheets.Count == 0)
                {
                    throw new InvalidOperationException("No worksheets found in the source file.");
                }

                // Assume we're working with the first worksheet for simplicity
                ExcelWorksheet sourceWorksheet = sourceWorkbook.Worksheets[1];

                // Create a new package for the target file
                using (var targetPackage = new ExcelPackage())
                {
                    ExcelWorkbook targetWorkbook = targetPackage.Workbook;
                    ExcelWorksheet targetWorksheet = targetWorkbook.Worksheets.Add("Sheet1"); // Name of the sheet in the target file

                    // Copy the header row (assuming you want headers, adjust if not needed)
                    // This example assumes headers are in row 1, adjust `rowStart` if different
                    if (rowStart > 1)
                    {
                        CopyRow(sourceWorksheet, targetWorksheet, 1, 1, false); // false: Don't merge cells for header
                    }

                    // Copy the specified rows with cell merging
                    for (int i = rowStart; i <= rowEnd; i++)
                    {
                        CopyRow(sourceWorksheet, targetWorksheet, i, i - (rowStart - 1) + (rowStart > 1 ? 1 : 0), true); // true: Merge cells for data rows
                    }

                    // Save the new file
                    FileStream targetStream = File.Create(targetFilePath);
                    targetPackage.SaveAs(targetStream);
                    targetStream.Close();
                }
            }
        }

        private void CopyRow(ExcelWorksheet sourceWorksheet, ExcelWorksheet targetWorksheet, int sourceRow, int targetRow, bool mergeCells)
        {
            int columnsCount = sourceWorksheet.Dimension.End.Column;
            for (int i = 1; i <= columnsCount; i++)
            {
                // Copy cell value
                targetWorksheet.Cells[targetRow, i].Value = sourceWorksheet.Cells[sourceRow, i].Value;

                // Copy formatting
                CopyCellFormatting(sourceWorksheet.Cells[sourceRow, i], targetWorksheet.Cells[targetRow, i]);
            }

            if (mergeCells)
            {
                // Merge cells in the target row (from column 1 to the last column with data)
                targetWorksheet.Cells[targetRow, 1, targetRow, columnsCount].Merge = true;
            }
        }

        private void CopyCellFormatting(ExcelRange sourceCell, ExcelRange targetCell)
        {
              

            targetCell.Style.Numberformat = sourceCell.Style.Numberformat;
            targetCell.Style.Font.Name = sourceCell.Style.Font.Name;
            targetCell.Style.Font.Size = sourceCell.Style.Font.Size;
            targetCell.Style.Font.Bold = sourceCell.Style.Font.Bold;
            targetCell.Style.Font.Italic = sourceCell.Style.Font.Italic;
            targetCell.Style.HorizontalAlignment = sourceCell.Style.HorizontalAlignment;
            targetCell.Style.VerticalAlignment = sourceCell.Style.VerticalAlignment;

            // Simplified border settings, directly assigning ExcelColor
            targetCell.Style.Border.Top.Style = sourceCell.Style.Border.Top.Style;
            //targetCell.Style.Border.Top.Color = sourceCell.Style.Border.Top.Color; // Direct assignment, no conversion needed

            targetCell.Style.Border.Bottom.Style = sourceCell.Style.Border.Bottom.Style;
            //targetCell.Style.Border.Bottom.Color = sourceCell.Style.Border.Bottom.Color;

            targetCell.Style.Border.Left.Style = sourceCell.Style.Border.Left.Style;
            //targetCell.Style.Border.Left.Color = sourceCell.Style.Border.Left.Color;

            targetCell.Style.Border.Right.Style = sourceCell.Style.Border.Right.Style;
            //targetCell.Style.Border.Right.Color = sourceCell.Style.Border.Right.Color;

            // For Fill (background) color, similar direct assignment
            targetCell.Style.Fill.PatternType = sourceCell.Style.Fill.PatternType;
            //targetCell.Style.Fill.BackgroundColor = sourceCell.Style.Fill.BackgroundColor; // Direct assignment
        }
    }
}
