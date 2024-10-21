using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;

namespace UnionCompl
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public MainWindow()
        {
            InitializeComponent();
            file_names = new List<string>();
        }

        private List<string> file_names;

        private void select_files_button_Click(object sender, RoutedEventArgs e)
        {
            file_names.Clear();
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();

            openFileDlg.Filter = "XLSX files (*.xlsx)|*.xlsx|XLS files (*.xls)|*.xls";
            openFileDlg.Title = "Выберите файлы комплектации";
            openFileDlg.FilterIndex = 0;
            openFileDlg.Multiselect = true;
            int a = file_names.Count;

            // Launch OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = openFileDlg.ShowDialog();

            // Get the selected file name and display in a TextBox.
            // Load content of file in a TextBlock
            if (result == true)
            {
                //file_names = openFileDlg.FileNames;
                if (openFileDlg.FileNames.Length > 0)
                {
                    foreach (string Item in openFileDlg.FileNames)
                    {
                        file_names.Add(Item);
                    }
                }
                //string files_1 = "";
                //foreach (string Item in file_names) files_1 += Item + "; ";
                //MessageBox.Show("Выбрали файлы: " + files_1);
                foreach (string Item in file_names)
                { 
                    if (!loaded_files_list_view.Items.Contains(Item))
                        loaded_files_list_view.Items.Add(Item);
                }

            }

        }

        private void clear_files_list_button_Click(object sender, RoutedEventArgs e)
        {
            loaded_files_list_view.Items.Clear();
        }

        private void select_path_button_Click(object sender, RoutedEventArgs e)
        {
            DateTime date = DateTime.Now;

            //string file_to_save = "";
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            openFileDlg.FileName = "Групповая комплектация " + date.Year.ToString() + date.Month.ToString() +
                date.Day.ToString() + ".xlsx";
            openFileDlg.Filter = "XLSX files (*.xlsx)|*.xlsx|XLS files (*.xls)|*.xls";
            openFileDlg.Title = "Выберите путь для сохранения файла";
            openFileDlg.FilterIndex = 0;
            openFileDlg.CheckFileExists = false;

            // Launch OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = openFileDlg.ShowDialog();

            // Get the selected file name and display in a TextBox.
            // Load content of file in a TextBlock
            if (result == true)
            {
                file_export_name.Text = openFileDlg.FileName;
            }
        }

        private void save_complect_button_Click(object sender, RoutedEventArgs e)
        {
            log_list_view.Items.Clear();
            
            // проверяем на наличие файлов в списке
            if (loaded_files_list_view.Items.Count == 0)
                return;

            // получаем данные о ширине столбцов, для того, чтобы корректно вывести данные первого файла

            float total_weight = 0f;
            float total_volume = 0f;

            // Загружаем данные в all_detail
            Dictionary<string, Dictionary<string, int>> all_detail = new Dictionary<string, Dictionary<string, int>>();
            foreach (string item in loaded_files_list_view.Items)
            { 
                Dictionary<string, Dictionary<string, int>> components = new Dictionary<string, Dictionary<string, int>>();
                ComplReader reader = new ComplReader(item);
                reader.read_file();
                components = reader.Elements;
                all_detail = ComplReader.merge_dict(all_detail, components);
                total_volume += reader.total_volume;
                total_weight += reader.total_weight;
            }

            // выводим загруженные данные в log_list_view
            foreach (string item in all_detail.Keys.OrderBy(x => x).ToList())
            { 
                log_list_view.Items.Add(item.ToString());
                List<string> components = all_detail[item].Keys.OrderBy(x=>x).ToList();                
                foreach (string name in components)
                { 
                    log_list_view.Items.Add(name + ": " + all_detail[item][name]);
                }
            }

            // добавляем логи с общим весом и объемом
            log_list_view.Items.Add("Общий вес: " + total_weight);
            log_list_view.Items.Add("Общий объем: " + total_volume);

            // клонируем ширину колонок в целевой файл
            ComplReader.CloneColumnWidths(loaded_files_list_view.Items[0].ToString(), file_export_name.Text);
            // копируем текст
            ComplWriter writer = new ComplWriter();
            writer.CopyRowsWithFormatting(loaded_files_list_view.Items[0].ToString(), file_export_name.Text, 1, 12);

        }
    }
}
