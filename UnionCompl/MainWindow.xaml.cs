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
            List<string> files = new List<string>();
            // Create OpenFileDialog
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
                string files_1 = "";
                foreach (string Item in file_names) files_1 += Item + "; ";
                MessageBox.Show("Выбрали файлы: " + files_1);

            }

        }
    }
}
