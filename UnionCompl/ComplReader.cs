using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
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

        public string file_name;

        public ComplReader()
        {
            Elements = new Dictionary<string, Dictionary<string, int>>();
            file_name = "";
        }

        public ComplReader(string file_name)
        {
            Elements = new Dictionary<string, Dictionary<string, int>>();
            this.file_name = file_name;
            //read_file();
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

                    for (int row = 12; row <= row_count; row++)
                    {
                        string directory = "";
                        Dictionary<int, string> row_dict = new Dictionary<int, string>();
                        for (int col = 1; col <= column_count; col++)
                        {
                            if (worksheet.Cells[row, col].Value != null)
                                row_dict.Add(col, worksheet.Cells[row, col].Value.ToString().Trim());
                            //list.Add (" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value.ToString().Trim());
                        }

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
                            // если радела еще нет
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
                if (row[1] == "")
                    if (!is_string_int(row[2]))
                        return new Tuple<bool, string>(true, row[2]);
            // Если первый символ цифра
            if (is_string_int(row[1]))
                return new Tuple<bool, string>(false, row[1]);

            return new Tuple<bool, string>(false, "false");
        }

        private bool is_string_int(string str)
        {
            if (int.TryParse(str, out int number))
                return true;
            else
                return false;
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

    }
}
