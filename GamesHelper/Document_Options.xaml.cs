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
using System.Windows.Shapes;

namespace GamesHelper
{
    /// <summary>
    /// Логика взаимодействия для Document_Options.xaml
    /// </summary>
    public partial class Document_Options : Window
    {
        private bool Exit = false;
        public Document_Options()
        {
            InitializeComponent();
        }

        private void Create_WORD_File_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            if (main != null)
            {
                int i = 0;
                if (For_Whom.IsChecked.Value)
                {
                    main.DocumentOptions[0] = "для родителей";
                }
                else
                {
                    main.DocumentOptions[0] = "для воспитателей";
                }

                if (Date.SelectedDate != null)
                {
                    main.DocumentOptions[1] = Convert.ToDateTime(Date.SelectedDate).ToString("d");
                    i++;
                }
                else
                {
                    MessageBox.Show("Выберите дату");
                }

                if (FIO.Text.Length > 5)
                {
                    main.DocumentOptions[2] = FIO.Text;
                    i++;
                }
                else
                {
                    MessageBox.Show("Введите ФИО");
                }

                if (Group.Text.Length > 1)
                {
                    main.DocumentOptions[3] = Group.Text;
                    i++;
                }
                else
                {
                    MessageBox.Show("Введите Группу");
                }

                if (Kindergarten.Text.Length > 1)
                {
                    main.DocumentOptions[4] = Kindergarten.Text;
                    i++;
                }
                else
                {
                    MessageBox.Show("Введите Сад");
                }

                if (i==4)
                {
                    Exit = true;
                    Create_WORD_File.IsEnabled = true;
                    this.Close();
                }
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (Exit)
            {
                MainWindow main = this.Owner as MainWindow;
                if (main != null)
                {
                    main.CreateDocumentWithOptions();
                }
                e.Cancel = false;
            }
            else
                e.Cancel = true;
        }
    }
}
