using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Automation.Peers;

namespace UnionCompl
{
    internal class ComplReader
    {
        /// <summary>
        /// Словарь с элементами комплектации
        /// </summary>
        public Dictionary<string, Dictionary<string, int>> Elements;

        /// <summary>
        /// Имя файла Excel с комплектацией
        /// </summary>
        public string file_name;
        
        /// <summary>
        /// Общий вес
        /// </summary>
        public float total_weight;

        /// <summary>
        /// Общий объем
        /// </summary>
        public float total_volume;

        /// <summary>
        /// Номер комплектаций
        /// </summary>
        public string compl_number;

        /// <summary>
        /// Наименование проекта
        /// </summary>
        public string prj_name;

        public ComplReader()
        {
            Elements = new Dictionary<string, Dictionary<string, int>>();
            file_name = "";
            total_weight = 0f;
            total_volume = 0f;
            compl_number = "";
            prj_name = "";
        }

        /// <summary>
        /// Конструктор
        /// </summary>
        /// <param name="file_name">имя файла excel</param>
        public ComplReader(string file_name)
        {
            Elements = new Dictionary<string, Dictionary<string, int>>();
            this.file_name = file_name;            
            total_weight = 0f;
            total_volume = 0f;
        }

        /// <summary>
        /// Читает файл и заполняет Elements
        /// </summary>
        public void read_file()
        {
            List<string> list = new List<string>();
            FileInfo existing_file = new FileInfo(file_name);
            if (existing_file.Exists)
            {
                using (ExcelPackage package = new ExcelPackage(existing_file))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                    int column_count = worksheet.Dimension.End.Column;
                    int row_count = worksheet.Dimension.End.Row;
                    string directory = "";

                    // get the compl_number
                    compl_number = CutStringAfterSymbol(worksheet.Cells[4, 1].Value.ToString(), "№");

                    // get the prj_name
                    if (worksheet.Cells[5, 3].Value !=null)
                        prj_name = worksheet.Cells[5, 3].Value.ToString();

                    for (int row = 12; row <= row_count; row++)
                    {
                        
                        Dictionary<int, string> row_dict = new Dictionary<int, string>();
                        for (int col = 1; col <= column_count; col++)
                        {
                            if (worksheet.Cells[row, col].Value != null)
                                row_dict.Add(col, worksheet.Cells[row, col].Value.ToString().Trim());
                            //list.Add (" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value.ToString().Trim());
                        }
                        if (row_dict.Count == 0)
                            continue;

                        Tuple<bool, string> check_dir = check_directory(row_dict);
                        if (check_dir.Item1)
                        {
                            directory = check_dir.Item2;
                        }
                        Tuple<bool, string, int> check_pos = check_position(row_dict);
                        if (check_pos.Item1)
                        {
                            // если раздел уже есть
                            if (Elements.ContainsKey(directory))
                            {
                                // если позиция уже есть в разделе
                                if (Elements[directory].ContainsKey(check_pos.Item2))
                                {
                                    // прибавляем значение
                                    Elements[directory][check_pos.Item2] += check_pos.Item3;
                                }
                                // если позиции еще нет в разделе
                                else
                                {
                                    Elements[directory].Add(check_pos.Item2, check_pos.Item3);
                                }
                            }
                            // если раздела еще нет
                            else
                            {
                                // создаем позицию
                                Dictionary<string, int> to_add = new Dictionary<string, int>
                                {
                                    { check_pos.Item2, check_pos.Item3 }
                                };
                                // добавляем ключ - раздел, значение - позиция
                                Elements.Add(directory, to_add );
                            }
                        }
                        
                        Tuple<bool,float> check_wei = cheсk_weight(row_dict);
                        if (check_wei.Item1) 
                        {
                            total_weight += check_wei.Item2;
                        }

                        Tuple<bool, float> check_vol = cheсk_volume(row_dict);
                        if (check_vol.Item1)
                        {
                            total_volume += check_vol.Item2;
                        }

                    }
                }
            }
        }


        /// <summary>
        /// Проверяет строку на соответсвие позиции комплектации
        /// Основные признаки позиции:
        ///     первый столбец - всегда целое число
        ///     второй и другие столбцы кроме последнего - текст
        ///     послежний снова - целое число
        /// </summary>
        /// <param name="row">Словарь, ключ - номер столбца, значение - текст в ячейке</param>
        /// <returns>true если позиция, "позиция", количество </returns>
        private Tuple<bool, string, int> check_position(Dictionary<int, string> row)
        {
            if (is_string_int(row[1]))
            {
                // пока количество сделаем жесткий 9 столбец
                // отсортируем по возрастанию
                List<int> column_numbers = row.Keys.OrderBy(x => x).ToList();
                string position_text = "";
                
                //int last_position = column_numbers.Max();
                // пока оставим жестко позицию 9
                int last_position = 9;
                if (!is_string_int(row[last_position]))
                    return new Tuple<bool, string, int>(false, "false", 0);
                foreach (int item in column_numbers)
                {
                    if (item != 1 && item !=last_position)
                    { 
                        position_text += row[item];
                    }
                }
                
                // считываем количество
                int amount = 0;
                if (int.TryParse(row[last_position], out int number))
                    amount = number;
                else
                    return new Tuple<bool, string, int>(false, "false", 0);
                
                return new Tuple<bool, string, int>(true, position_text, amount);
            }
            return new Tuple<bool, string, int>(false, "false", 0);
        }




        /// <summary>
        /// проверяет является ли строка заголовком группы деталей
        /// </summary>
        /// <param name="row">Словарь, ключ - номер столбца, значение - текст в ячейке</param>
        /// <returns>true - если заголовок, "Наименование заголовка" </returns>
        private Tuple<bool, string> check_directory(Dictionary<int, string> row)
        {
            // если поле пустое
            if (row.Count == 0)
                return new Tuple<bool, string>(false, "");
            // если значимых ячеек больше чем одна
            if (row.Count == 2)
                if (row[1] == "" && row.ContainsKey(2))
                    if (!is_string_int(row[2]))
                        return new Tuple<bool, string>(true, row[2]);
            // Если первый символ цифра
            if (is_string_int(row[1]))
                return new Tuple<bool, string>(false, row[1]);
            return new Tuple<bool, string>(false, "false");
        }

        /// <summary>
        /// Проверяет является ли строка целым числом
        /// </summary>
        /// <param name="str">Строка</param>
        /// <returns>true если является, false - если нет</returns>
        private bool is_string_int(string str)
        {
            if (int.TryParse(str, out int number))
                return true;
            else
                return false;
        }

        /// <summary>
        /// проверяет является ли строка весом
        /// </summary>
        /// <param name="row">Словарь, ключ - номер столбца, значение - текст в ячейке</param>
        /// <returns>кортеж true, если вес обнаружен и значение, false если нет</returns>
        private Tuple<bool, float> cheсk_weight(Dictionary<int, string> row)
        {
            if (row.Count == 0)
                return new Tuple<bool, float>(false, 0.0f);
            if (row.Count == 2)
                if (row.ContainsKey(2) && row[2].Length > 10)
                {
                    string a = row[2].Substring(0, 10);
                    if (row[2].Substring(0, 10) == "Общий вес:")
                    {
                        // распознаем цифры
                        float? NullableFloat = ExtractNumberFromString(row[2]);
                        if (NullableFloat.HasValue)
                        {                                                        
                            return new Tuple<bool, float>(true, NullableFloat.Value);
                        }
                    }                           
                }
            return new Tuple<bool, float>(false, 0.0f);
        }


        /// <summary>
        /// проверяет является ли строка весом
        /// </summary>
        /// <param name="row">Словарь, ключ - номер столбца, значение - текст в ячейке</param>
        /// <returns>кортеж true, если вес обнаружен и значение, false если нет</returns>
        private Tuple<bool, float> cheсk_volume(Dictionary<int, string> row)
        {
            if (row.Count == 0)
                return new Tuple<bool, float>(false, 0.0f);
            if (row.Count == 2)
                if (row.ContainsKey(2) && row[2].Length > 10)
                {
                    string a = row[2].Substring(0, 12);
                    if (row[2].Substring(0, 12) == "Общий объем:")
                    {
                        // распознаем цифры
                        float? NullableFloat = ExtractNumberFromString(row[2]);
                        if (NullableFloat.HasValue)
                        {
                            return new Tuple<bool, float>(true, NullableFloat.Value);
                        }
                    }
                }
            return new Tuple<bool, float>(false, 0.0f);

        }


        /// <summary>
        /// Соединяет словари словарей между собой в dict_1
        /// </summary>
        /// <param name="dict_1"></param>
        /// <param name="dict_2"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, int>> merge_dict(Dictionary<string, Dictionary<string, int>> dict_1, Dictionary<string, Dictionary<string, int>> dict_2)
        {
            foreach (string directory in dict_2.Keys)
            {
                // если есть раздел в dict1
                if (dict_1.ContainsKey(directory))
                {
                    // перебираем   словарь
                    foreach (string item in dict_2[directory].Keys)
                    {
                        // если в словаре уже есть позиция
                        if (dict_1[directory].ContainsKey(item))
                        {
                            dict_1[directory][item] += dict_2[directory][item];
                        }
                        // если в словаре нет позиции
                        else
                        {
                            dict_1[directory].Add(item, dict_2[directory][item]);
                        }
                    }
                }
                // если раздела нет в dict1
                else
                {
                    dict_1.Add(directory, dict_2[directory]);
                }
            }
            return dict_1;
        }

        /// <summary>
        /// Ищет в текстовой строке число и возвращает их во float?
        /// </summary>
        /// <param name="input">Текстовая строка</param>
        /// <returns></returns>
        public static float? ExtractNumberFromString(string input)
        {
            // Regular expression pattern to match both integers and floats
            // with either '.' or ',' as the decimal separator, considering
            // the decimal separator might be followed by more digits
            var numberPattern = new Regex(@"\d+(?:[.,]\d+)?", RegexOptions.CultureInvariant);

            // Find the first occurrence of the pattern in the string
            var match = numberPattern.Match(input);

            // If a match is found, attempt to convert it to a float
            if (match.Success)
            {
                var numberString = match.Value;
                // Replace ',' with '.' if ',' was used as the decimal separator
                // to facilitate conversion to float
                if (numberString.Contains(","))
                    numberString = numberString.Replace(",", ".");

                // Attempt to parse the string to a float
                CultureInfo userCulture = CultureInfo.InvariantCulture; // or get from user preferences
                if (float.TryParse(numberString, NumberStyles.Float, userCulture, out float number))
                    return number;
                else
                    return null;
                    //throw new FormatException($"Failed to convert '{numberString}' to a float.");
            }

            // If no match is found or conversion fails, return null
            return null;
        }


        /// <summary>
        /// Cuts the input string at the specified symbol and returns the part after the symbol.
        /// If the symbol is not found, an empty string is returned.
        /// </summary>
        /// <param name="originalString">The string to be cut.</param>
        /// <param name="cutSymbol">The symbol at which to cut the string.</param>
        /// <returns>The part of the string after the cut symbol, or an empty string if not found.</returns>
        public static string CutStringAfterSymbol(string originalString, string cutSymbol)
        {
            if (string.IsNullOrEmpty(originalString) || string.IsNullOrEmpty(cutSymbol))
            {
                return string.Empty; // Return empty if either string is null or empty
            }

            int symbolIndex = originalString.IndexOf(cutSymbol);

            // If the symbol is not found, IndexOf returns -1, so return an empty string
            if (symbolIndex == -1)
            {
                return string.Empty;
            }

            // We add the length of the cutSymbol to the index to start the substring AFTER the symbol
            return originalString.Substring(symbolIndex + cutSymbol.Length);
        }



        /// <summary>
        /// Клонирует ширину колонок из sourceFilePath в targetFilePath
        /// </summary>
        /// <param name="sourceFilePath">Исходный файо</param>
        /// <param name="targetFilePath">Целевой файл</param>
        /// <exception cref="FileNotFoundException"></exception>
        /// <exception cref="InvalidOperationException"></exception>
        public static void CloneColumnWidths(string sourceFilePath, string targetFilePath)
        {
            // Check if source file exists
            if (!File.Exists(sourceFilePath))
            {
                throw new FileNotFoundException("Не найден исходный файл", sourceFilePath);
            }

            // Load the source Excel file
            using (var package = new ExcelPackage(new FileInfo(sourceFilePath)))
            {
                var sourceWorkbook = package.Workbook;
                if (sourceWorkbook.Worksheets.Count < 1)
                {
                    throw new InvalidOperationException("No worksheets found in the source file.");
                }

                // Assume we're working with the first worksheet for simplicity
                var sourceSheet = sourceWorkbook.Worksheets[1];

                // Create a new Excel package for the target file
                using (var targetPackage = new ExcelPackage())
                {
                    var targetWorkbook = targetPackage.Workbook;
                    var targetSheet = targetWorkbook.Worksheets.Add("Sheet1"); // Name of the new sheet

                    // Iterate through the columns of the source sheet to set widths in the target sheet
                    for (int column = 1; column <= sourceSheet.Dimension.End.Column; column++)
                    {
                        // Get the column width from the source sheet
                        var columnWidth = sourceSheet.Column(column).Width;

                        // If the width is not set (i.e., it's the default), you might want to handle this differently
                        // depending on your requirements. Here, we just apply it as is.
                        targetSheet.Column(column).Width = columnWidth;
                    }

                    // Save the new Excel file
                    var targetFileInfo = new FileInfo(targetFilePath);
                    targetPackage.SaveAs(targetFileInfo);
                }
            }
        }

    }
}
