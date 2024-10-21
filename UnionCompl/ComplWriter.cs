using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Text.RegularExpressions;
using System.IO.Packaging;

namespace UnionCompl
{
    internal class ComplWriter
    {
        public ComplWriter()
        {

        }

        public static void write_compl(string sourceFilePath, string targetFilePath, int rowStart, int rowEnd, Dictionary<string, Dictionary<string, int>> elements, float total_weight = 0, float total_volume = 0, string compl_numbers ="", string prj_names ="")
        {
            if (rowStart < 1 || rowEnd < 1 || rowEnd < rowStart)
            {
                throw new ArgumentException("Invalid row numbers. Both must be positive and rowEnd must be greater than or equal to rowStart.");
            }

            if (!File.Exists(sourceFilePath))
            {
                throw new FileNotFoundException("Source file not found.", sourceFilePath);
            }

            const string DefaultFont = "Arial";
            
            using (var package = new ExcelPackage(new FileInfo(sourceFilePath)))
            {
                ExcelWorkbook sourceWorkbook = package.Workbook;
                if (sourceWorkbook.Worksheets.Count == 0)
                {
                    throw new InvalidOperationException("No worksheets found in the source file.");
                }

                // Assume we're working with the first worksheet for simplicity
                ExcelWorksheet sourceWorksheet = sourceWorkbook.Worksheets[1];

                using (var targetPackage = new ExcelPackage())
                {
                    ExcelWorkbook targetWorkbook = targetPackage.Workbook;
                    ExcelWorksheet targetWorksheet = targetWorkbook.Worksheets.Add("Sheet1"); // Name of the sheet in the target file

                    int totalRows = sourceWorksheet.Dimension.End.Row;
                    int totalCols = sourceWorksheet.Dimension.End.Column;

                    for (int column = 1; column <= sourceWorksheet.Dimension.End.Column; column++)
                    {
                        // Get the column width from the source sheet
                        var columnWidth = sourceWorksheet.Column(column).Width;

                        // If the width is not set (i.e., it's the default), you might want to handle this differently
                        // depending on your requirements. Here, we just apply it as is.
                        targetWorksheet.Column(column).Width = columnWidth;
                    }


                    sourceWorksheet.Cells[rowStart, 1, rowEnd, totalCols].Copy(targetWorksheet.Cells[rowStart, 1, rowEnd, totalCols]);
                    // Save the new file
                    
                    

                    int current_row = 13;
                    int position_counter = 1;

                    foreach (string directory in elements.Keys.OrderBy(x => x).ToList())
                    {
                        targetWorksheet.Cells[current_row, 1].Value = "";

                        targetWorksheet.Cells[current_row, 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        targetWorksheet.Cells[current_row, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        targetWorksheet.Cells[current_row, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        targetWorksheet.Cells[current_row, 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                        targetWorksheet.Cells[current_row, 2].Value = directory;
                        targetWorksheet.Cells[current_row, 2].Style.Font.Size = 11;
                        targetWorksheet.Cells[current_row, 2].Style.Font.Bold = true;
                        targetWorksheet.Cells[current_row, 2].Style.Font.Name = DefaultFont;                        

                        targetWorksheet.Cells[current_row, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        targetWorksheet.Cells[current_row, 2, current_row, 9].Merge = true;
                        targetWorksheet.Cells[current_row, 2, current_row, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        targetWorksheet.Cells[current_row, 2, current_row, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        targetWorksheet.Cells[current_row, 2, current_row, 9].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        targetWorksheet.Cells[current_row, 2, current_row, 9].Style.Border.Right.Style = ExcelBorderStyle.Thin;


                        current_row += 1;
                        foreach (string position in elements[directory].Keys.OrderBy(x => x).ToList())
                        {
                            targetWorksheet.Cells[current_row, 1].Style.Font.Name = DefaultFont;
                            targetWorksheet.Cells[current_row, 1].Style.Font.Size = 10;
                            targetWorksheet.Cells[current_row, 1].Value = position_counter;
                            targetWorksheet.Cells[current_row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            targetWorksheet.Cells[current_row, 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            targetWorksheet.Cells[current_row, 1].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            targetWorksheet.Cells[current_row, 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            targetWorksheet.Cells[current_row, 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;


                            targetWorksheet.Cells[current_row, 2].Style.Font.Name = DefaultFont;
                            targetWorksheet.Cells[current_row, 2].Style.Font.Size = 10;
                            targetWorksheet.Cells[current_row, 2].Value = position;
                            targetWorksheet.Cells[current_row, 2, current_row, 8].Merge = true;

                            targetWorksheet.Cells[current_row, 2, current_row, 8].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            targetWorksheet.Cells[current_row, 2, current_row, 8].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            targetWorksheet.Cells[current_row, 2, current_row, 8].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            targetWorksheet.Cells[current_row, 2, current_row, 8].Style.Border.Right.Style = ExcelBorderStyle.Thin;


                            targetWorksheet.Cells[current_row, 9].Style.Font.Name = DefaultFont;
                            targetWorksheet.Cells[current_row, 9].Style.Font.Size = 10;
                            targetWorksheet.Cells[current_row, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            targetWorksheet.Cells[current_row, 9].Value = elements[directory][position];

                            targetWorksheet.Cells[current_row, 9].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                            targetWorksheet.Cells[current_row, 9].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                            targetWorksheet.Cells[current_row, 9].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                            targetWorksheet.Cells[current_row, 9].Style.Border.Right.Style = ExcelBorderStyle.Thin;


                            current_row += 1;
                            position_counter += 1;
                        }
                    }

                    current_row += 1;
                    targetWorksheet.Cells[current_row, 2].Style.Font.Name = DefaultFont;
                    targetWorksheet.Cells[current_row, 2].Style.Font.Size = 10;
                    targetWorksheet.Cells[current_row, 2].Style.Font.Bold = true;
                    targetWorksheet.Cells[current_row, 2].Value = "Общий вес: " + total_weight + " кг.";
                    current_row += 1;
                    targetWorksheet.Cells[current_row, 2].Style.Font.Name = DefaultFont;
                    targetWorksheet.Cells[current_row, 2].Style.Font.Size = 10;
                    targetWorksheet.Cells[current_row, 2].Style.Font.Bold = true;
                    targetWorksheet.Cells[current_row, 2].Value = "Общий объем:: " + total_volume + " куб.м.";

                    targetWorksheet.Cells[4, 1].Value = "Групповая комплектация №" + compl_numbers;
                    targetWorksheet.Cells[5, 3].Value = prj_names;

                    FileStream targetStream = File.Create(targetFilePath);
                    targetPackage.SaveAs(targetStream);
                    targetStream.Close();
                }
            }

        }

    }
}
