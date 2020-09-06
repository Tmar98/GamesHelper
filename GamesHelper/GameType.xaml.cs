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
    /// Логика взаимодействия для GameType.xaml
    /// </summary>
    public partial class GameType : Window
    {
        private bool Exit = false;
        public GameType()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            MainWindow main = this.Owner as MainWindow;
            if (main != null)
            {
                if (One.IsChecked.Value)
                {
                    main.GameType_String = "и";
                    Exit = true;
                    this.Close();
                }
                else if (Groupe.IsChecked.Value)
                {
                    main.GameType_String = "гр";
                    Exit = true;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Выберите тип игры");
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
                    main.GameTypeXamlClosed();
                }
                e.Cancel = false;
            }
            else
                e.Cancel = true;
            
        }
    }
}
