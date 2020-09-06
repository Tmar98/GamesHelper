using System;
using System.Windows;
using System.IO;
using Microsoft.Office.Interop.Word;
using Window = System.Windows.Window;
using Table = Microsoft.Office.Interop.Word.Table;

namespace GamesHelper
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string FileName1 = @"WorkDirections.txt";
        private string FileName2 = @"Tasks.txt";
        private string FileName3 = "Games.xml";
        private string FileName4 = @"GameBinding.txt";
        public string GameType_String = "";

        public MainWindow()
        {
            InitializeComponent();
           List_WorkDirection list_WorkDirection = new List_WorkDirection(FileName1);
            List_Direction_UI.ItemsSource = list_WorkDirection;
            List_Direction_UI.DisplayMemberPath = "Work";
        }

        private void Butt_WorkDirection_Click(object sender, RoutedEventArgs e) //показываются задачи на основе выбранного направления
        {
            if (List_Direction_UI.SelectedItem != null)
            {
                List_Direction_UI.Visibility = Visibility.Hidden;
                List_Task_UI.Visibility = Visibility.Visible;
                Butt_WorkDirection.Visibility = Visibility.Hidden;
                Butt_Task.Visibility = Visibility.Visible;

                List_Tasks list_Tasks = new List_Tasks(FileName2, List_Direction_UI.SelectedIndex + 1);
                List_Task_UI.ItemsSource = list_Tasks;
            }
        }

        private void Butt_Task_Click(object sender, RoutedEventArgs e)//показываются игры на основе выбранной задачи
        {
            GameType gameType = new GameType();
            gameType.Owner = this;
            gameType.Show();
            Butt_Task.IsEnabled = false;
            string are = "";
            
        }
        public void GameTypeXamlClosed ()
        {
            Gamesto_List();
        }
        private void Gamesto_List()
        {
                List_Task_UI.Visibility = Visibility.Hidden;
                List_Game_UI.Visibility = Visibility.Visible;
                Butt_Task.Visibility = Visibility.Hidden;
                Butt_Games.Visibility = Visibility.Visible;
                Butt_GametoList.Visibility = Visibility.Visible;

                List_Games list_Games = new List_Games(FileName3);
                var t = List_Task_UI.SelectedItem;
                var binding_List_Games = Read_Games_Bindings(FileName4, list_Games, (Task)List_Task_UI.SelectedItem);

                List_Game_UI.ItemsSource = binding_List_Games;
        }
        private List_Binding_Game Read_Games_Bindings(string fileName, List_Games list_Games,Task task_id)
        {
            List_Binding_Game list_binding_Game = new List_Binding_Game();//новый класс лист с нумерацией
            if (File.Exists(fileName))
            {
                StreamReader sr = new StreamReader(fileName);
                string s0;
                while ((s0 = sr.ReadLine()) != null)
                {
                    string[] ar = s0.Split(' ');
                    if(Convert.ToInt32( ar[0]) == task_id.Id)
                    {
                        if ((list_Games[Convert.ToInt32(ar[1]) - 1].Id == Convert.ToInt32(ar[1]) )&( list_Games[Convert.ToInt32(ar[1]) - 1].GameType.Contains(GameType_String)))//сходятся ли id игры
                        {
                            Binding_Game binding_game = new Binding_Game(list_binding_Game.Count + 1,list_Games[Convert.ToInt32(ar[1]) - 1]);//новый класс игры с нумерацией + добавление нумерации
                            list_binding_Game.Add(binding_game);//добавление в лист
                        }
                    }
                    
                }
                sr.Close();
                sr.Dispose();
            }
            return list_binding_Game;
        }

        private void CreateDocument()
        {
            try
            {
                //Create an instance for word app  
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

                //Set animation status for word application  
                winword.ShowAnimation = false;
                
                //Set status for word application is to be visible or not.  
                winword.Visible = false;

                

                //Create a missing variable for missing value  
                object missing = System.Reflection.Missing.Value;

                //Create a new document  
                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add();

                document.PageSetup.Orientation = WdOrientation.wdOrientLandscape;//альбомная ориентация

                //Add header into the document  
                foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                {
                    //Get the header range and add the header details.  
                    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                    headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                    headerRange.Font.Size = 10;
                    headerRange.Text = "Header text goes here";
                }

                //Add the footers into the document  
                foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
                {
                    //Get the footer range and add the footer details.  
                    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                    footerRange.Font.Size = 10;
                    footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    footerRange.Text = "Footer text goes here";
                }

                
                
                //adding text to document  
                document.Content.SetRange(0, 0);
                document.Content.Text = "This is test document " + Environment.NewLine;

                //Add paragraph with Heading 1 style  
                Microsoft.Office.Interop.Word.Paragraph para1 = document.Content.Paragraphs.Add();
                para1.Range.set_Style(WdBuiltinStyle.wdStyleHeading1);
                para1.Range.Text = "Para 1 text";
                para1.Range.InsertParagraphAfter();

                //Add paragraph with Heading 2 style  
                Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add(ref missing);

                para2.Range.Text = "Para 2 text";
                para2.Range.InsertParagraphAfter();

                //Create a 5X5 table and insert some dummy record  
                Table firstTable = document.Tables.Add(para1.Range, 5, 5, ref missing, ref missing);

                firstTable.Borders.Enable = 1;
                foreach (Row row in firstTable.Rows)
                {
                    foreach (Cell cell in row.Cells)
                    {
                        //Header row  
                        if (cell.RowIndex == 1)
                        {
                            cell.Range.Text = "Column " + cell.ColumnIndex.ToString();
                            cell.Range.Font.Bold = 1;
                            //other format properties goes here  
                            cell.Range.Font.Name = "verdana";
                            cell.Range.Font.Size = 10;
                            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                            cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                            //Center alignment for the Header cells  
                            cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                            cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                        }
                        //Data row  
                        else
                        {
                            cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                        }
                    }
                }

                

                //Save the document  
                object filename = @"F:\doc2.docx";
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                MessageBox.Show("Document created successfully !");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void Butt_GametoList_Click(object sender, RoutedEventArgs e)
        {
            CreateDocument();
        }
    }
    

    
}
