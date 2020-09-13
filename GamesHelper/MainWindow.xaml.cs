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
        public string[] DocumentOptions = new string[5];

        private string[,] Task_Massiv = new string[10, 5];
        private string[,] Game_Massiv = new string[10, 5];

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
                int i = 0;
                foreach(Task task in list_Tasks)
                {
                    task.List_Id = i + 1;
                    i++;
                }
                List_Task_UI.ItemsSource = list_Tasks;
            }
        }

        private void Butt_Task_Click(object sender, RoutedEventArgs e)//показываются игры на основе выбранной задачи
        {
            if (List_Task_UI.SelectedItem != null)
            {
                GameType gameType = new GameType();
                gameType.Owner = this;
                gameType.ShowDialog();
            }


        }
        public void GameTypeXamlClosed()
        {
            Gamesto_List();
        }
        private void Gamesto_List()
        {
            List_Task_UI.Visibility = Visibility.Hidden;
            List_Game_UI.Visibility = Visibility.Visible;
            Butt_Task.Visibility = Visibility.Hidden;
            Butt_Games.Visibility = Visibility.Visible;

            List_Games list_Games = new List_Games(FileName3);
            var t = List_Task_UI.SelectedItem;
            var binding_List_Games = Read_Games_Bindings(FileName4, list_Games, (Task)List_Task_UI.SelectedItem);

            List_Game_UI.ItemsSource = binding_List_Games;
        }
        private List_Binding_Game Read_Games_Bindings(string fileName, List_Games list_Games, Task task_id)
        {
            List_Binding_Game list_binding_Game = new List_Binding_Game();//новый класс лист с нумерацией
            if (File.Exists(fileName))
            {
                StreamReader sr = new StreamReader(fileName);
                string s0;
                while ((s0 = sr.ReadLine()) != null)
                {
                    string[] ar = s0.Split(' ');
                    if (Convert.ToInt32(ar[0]) == task_id.Id)
                    {
                        if ((list_Games[Convert.ToInt32(ar[1]) - 1].Id == Convert.ToInt32(ar[1])) & (list_Games[Convert.ToInt32(ar[1]) - 1].GameType.Contains(GameType_String)))//сходятся ли id игры
                        {
                            Binding_Game binding_game = new Binding_Game(list_binding_Game.Count + 1, list_Games[Convert.ToInt32(ar[1]) - 1]);//новый класс игры с нумерацией + добавление нумерации
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
                Document document = winword.Documents.Add();

                document.PageSetup.Orientation = WdOrientation.wdOrientLandscape;//альбомная ориентация

                //Add header into the document  
                //foreach (Microsoft.Office.Interop.Word.Section section in document.Sections)
                //{
                //    //Get the header range and add the header details.  
                //    Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //    headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                //    headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //    headerRange.Font.Bold = 1;
                //    headerRange.Font.Size = 14;
                //    headerRange.Text = "Домашнее задание "+DocumentOptions[0]+" на закрепление пройденного материала";
                //}

                ////Add the footers into the document  
                //foreach (Microsoft.Office.Interop.Word.Section wordSection in document.Sections)
                //{
                //    //Get the footer range and add the footer details.  
                //    Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //    footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                //    footerRange.Font.Size = 10;
                //    footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                //    footerRange.Text = "Footer text goes here";
                //}

                Paragraph para1 = document.Content.Paragraphs.Add();
                para1.Range.Font.Size = 14;
                para1.Range.Font.Bold = 1;
                para1.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para1.Range.Text = "Домашнее задание " + DocumentOptions[0] + " на закрепление пройденного материала" + Environment.NewLine;
                para1.Range.InsertParagraphAfter();

                //adding text to document  
                //document.Content.SetRange(0, 0);
                //document.Content.Text = "Дата " + DocumentOptions[1] + " Ф.И ребенка " + DocumentOptions[2] + " группа " + DocumentOptions[3] + " сад №" + DocumentOptions[4] + Environment.NewLine;

                //Add paragraph with Heading 1 style  
                Paragraph para2 = document.Content.Paragraphs.Add();
                para2.Range.Font.Size = 12;
                para2.Range.Text = "Дата " + DocumentOptions[1] + "     Ф.И ребенка " + DocumentOptions[2] + "      группа " + DocumentOptions[3] + "       сад №" + DocumentOptions[4] + Environment.NewLine;
                para2.Range.InsertParagraphAfter();

                ////Add paragraph with Heading 2 style  
                //Microsoft.Office.Interop.Word.Paragraph para2 = document.Content.Paragraphs.Add(ref missing);

                //para2.Range.Text = "Para 2 text";
                //para2.Range.InsertParagraphAfter();

                //Create a 5X5 table and insert some dummy record  
                Table firstTable = document.Tables.Add(para1.Range, 4, 6, ref missing, ref missing);

                firstTable.Rows[1].Cells[1].Range.Text = "Направления работы";
                firstTable.Rows[1].Cells[2].Range.Text = "Социализация (представления о себе и окружающем мире, игра)";
                firstTable.Rows[1].Cells[3].Range.Text = "Коммуникативная сфера(общение, речь)";
                firstTable.Rows[1].Cells[4].Range.Text = "Эмоционально – волевая сфера(эмоции, чувства, воля, мотивация)";
                firstTable.Rows[1].Cells[5].Range.Text = "Психические процессы(память, внимание, воображение, пространственное восприятие)";
                firstTable.Rows[1].Cells[6].Range.Text = "Сенсомоторная сфера(ощущения, восприятие, крупная и мелкая моторика)";


                firstTable.Rows[2].Cells[1].Range.Text = "Задачи";
                string str11="";
                for (int i = 0; i <10;i++)
                {
                    if (Task_Massiv[i, 0] != null)
                        str11 += Task_Massiv[i, 0] + "\n";
                    else
                        i = 10;
                }
                string str12 = "";
                for (int i = 0; i < 10; i++)
                {
                    if (Task_Massiv[i, 1] != null)
                        str12 += Task_Massiv[i, 1] + "\n";
                    else
                        i = 10;
                }
                string str13 = "";
                for (int i = 0; i < 10; i++)
                {
                    if (Task_Massiv[i, 2] != null)
                        str13 += Task_Massiv[i, 2] + "\n";
                    else
                        i = 10;
                }
                string str14 = "";
                for (int i = 0; i < 10; i++)
                {
                    if (Task_Massiv[i, 3] != null)
                        str14 += Task_Massiv[i, 3] + "\n";
                    else
                        i = 10;
                }
                string str15 = "";
                for (int i = 0; i < 10; i++)
                {
                    if (Task_Massiv[i, 4] != null)
                        str15 += Task_Massiv[i, 4] + "\n";
                    else
                        i = 10;
                }
                firstTable.Rows[2].Cells[2].Range.Text = str11; /*Task_Massiv[0,0]+"\n"+ Task_Massiv[1,0]+ "\n" + Task_Massiv[2,0]+ "\n" + Task_Massiv[3,0] + "\n" + Task_Massiv[4,0] + "\n" + Task_Massiv[5,0] + "\n" + Task_Massiv[6,0] + "\n" + Task_Massiv[7,0] + "\n" + Task_Massiv[8, 0] + "\n" + Task_Massiv[9, 0];*/
                firstTable.Rows[2].Cells[3].Range.Text = str12; /*Task_Massiv[0, 1] + "\n" + Task_Massiv[1, 1] + "\n" + Task_Massiv[2, 1] + "\n" + Task_Massiv[3, 1] + "\n" + Task_Massiv[4, 1] + "\n" + Task_Massiv[5, 1] + "\n" + Task_Massiv[6, 1] + "\n" + Task_Massiv[7, 1] + "\n" + Task_Massiv[8, 1] + "\n" + Task_Massiv[9, 1];*/
                firstTable.Rows[2].Cells[4].Range.Text = str13; /*Task_Massiv[0, 2] + "\n" + Task_Massiv[1, 2] + "\n" + Task_Massiv[2, 2] + "\n" + Task_Massiv[3, 2] + "\n" + Task_Massiv[4, 2] + "\n" + Task_Massiv[5, 2] + "\n" + Task_Massiv[6, 2] + "\n" + Task_Massiv[7, 2] + "\n" + Task_Massiv[8, 2] + "\n" + Task_Massiv[9, 2];*/
                firstTable.Rows[2].Cells[5].Range.Text = str14; /*Task_Massiv[0, 3] + "\n" + Task_Massiv[1, 3] + "\n" + Task_Massiv[2, 3] + "\n" + Task_Massiv[3, 3] + "\n" + Task_Massiv[4, 3] + "\n" + Task_Massiv[5, 3] + "\n" + Task_Massiv[6, 3] + "\n" + Task_Massiv[7, 3] + "\n" + Task_Massiv[8, 3] + "\n" + Task_Massiv[9, 3];*/
                firstTable.Rows[2].Cells[6].Range.Text = str15; /*Task_Massiv[0, 4] + "\n" + Task_Massiv[1, 4] + "\n" + Task_Massiv[2, 4] + "\n" + Task_Massiv[3, 4] + "\n" + Task_Massiv[4, 4] + "\n" + Task_Massiv[5, 4] + "\n" + Task_Massiv[6, 4] + "\n" + Task_Massiv[7, 4] + "\n" + Task_Massiv[8, 4] + "\n" + Task_Massiv[9, 4];*/


                string str21 = "";
                for (int i = 0; i < 10; i++)
                {
                    if (Game_Massiv[i, 0] != null)
                        str21 += Game_Massiv[i, 0] + "\n";
                    else
                        i = 10;
                }
                string str22 = "";
                for (int i = 0; i < 10; i++)
                {
                    if (Game_Massiv[i, 1] != null)
                        str22 += Game_Massiv[i, 1] + "\n";
                    else
                        i = 10;
                }
                string str23 = "";
                for (int i = 0; i < 10; i++)
                {
                    if (Game_Massiv[i, 2] != null)
                        str23 += Game_Massiv[i, 2] + "\n";
                    else
                        i = 10;
                }
                string str24 = "";
                for (int i = 0; i < 10; i++)
                {
                    if (Game_Massiv[i, 3] != null)
                        str24 += Game_Massiv[i, 3] + "\n";
                    else
                        i = 10;
                }
                string str25 = "";
                for (int i = 0; i < 10; i++)
                {
                    if (Game_Massiv[i, 4] != null)
                        str25 += Game_Massiv[i, 4] + "\n";
                    else
                        i = 10;
                }


                firstTable.Rows[3].Cells[1].Range.Text = "Игры, задания";
                firstTable.Rows[3].Cells[2].Range.Text =str21; /*Game_Massiv[0, 0] + "\n" + Game_Massiv[1, 0] + "\n" + Game_Massiv[2, 0] + "\n" + Game_Massiv[3, 0] + "\n" + Game_Massiv[4, 0] + "\n" + Game_Massiv[5, 0] + "\n" + Game_Massiv[6, 0] + "\n" + Game_Massiv[7, 0] + "\n" + Game_Massiv[8, 0] + "\n" + Game_Massiv[9, 0];*/
                firstTable.Rows[3].Cells[3].Range.Text =str22; /*Game_Massiv[0, 1] + "\n" + Game_Massiv[1, 1] + "\n" + Game_Massiv[2, 1] + "\n" + Game_Massiv[3, 1] + "\n" + Game_Massiv[4, 1] + "\n" + Game_Massiv[5, 1] + "\n" + Game_Massiv[6, 1] + "\n" + Game_Massiv[7, 1] + "\n" + Game_Massiv[8, 1] + "\n" + Game_Massiv[9, 1];*/
                firstTable.Rows[3].Cells[4].Range.Text =str23; /*Game_Massiv[0, 2] + "\n" + Game_Massiv[1, 2] + "\n" + Game_Massiv[2, 2] + "\n" + Game_Massiv[3, 2] + "\n" + Game_Massiv[4, 2] + "\n" + Game_Massiv[5, 2] + "\n" + Game_Massiv[6, 2] + "\n" + Game_Massiv[7, 2] + "\n" + Game_Massiv[8, 2] + "\n" + Game_Massiv[9, 2];*/
                firstTable.Rows[3].Cells[5].Range.Text =str24; /*Game_Massiv[0, 3] + "\n" + Game_Massiv[1, 3] + "\n" + Game_Massiv[2, 3] + "\n" + Game_Massiv[3, 3] + "\n" + Game_Massiv[4, 3] + "\n" + Game_Massiv[5, 3] + "\n" + Game_Massiv[6, 3] + "\n" + Game_Massiv[7, 3] + "\n" + Game_Massiv[8, 3] + "\n" + Game_Massiv[9, 3];*/
                firstTable.Rows[3].Cells[6].Range.Text =str25; /*Game_Massiv[0, 4] + "\n" + Game_Massiv[1, 4] + "\n" + Game_Massiv[2, 4] + "\n" + Game_Massiv[3, 4] + "\n" + Game_Massiv[4, 4] + "\n" + Game_Massiv[5, 4] + "\n" + Game_Massiv[6, 4] + "\n" + Game_Massiv[7, 4] + "\n" + Game_Massiv[8, 4] + "\n" + Game_Massiv[9, 4];*/



                firstTable.Rows[4].Height = 200;
                firstTable.Rows[4].Cells[1].Range.Text = "Результат, примечания* ";

                firstTable.Borders.Enable = 1;

                //foreach (Row row in firstTable.Rows)
                //{
                //    foreach (Cell cell in row.Cells)
                //    {


                //        //Header row  
                //        if (cell.RowIndex == 1)
                //        {
                //            cell.Range.Text = "Column " + cell.ColumnIndex.ToString();
                //            cell.Range.Font.Bold = 1;
                //            //other format properties goes here  
                //            cell.Range.Font.Name = "verdana";
                //            cell.Range.Font.Size = 10;
                //            //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                //            cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                //            //Center alignment for the Header cells  
                //            cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                //            cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                //        }
                //        //Data row  
                //        else
                //        {
                //            cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                //        }
                //    }
                //}


                Paragraph para3 = document.Content.Paragraphs.Add();
                para3.Range.Font.Size = 12;
                if (DocumentOptions[0] == "для родителей")
                {
                    para3.Range.Text = "* Заполняется родителями (выполняет с легкостью/ справляется/ возникают незначительные затруднения/ возникли значительные трудности/ отказ от выполнения задания/  не выполнили задания; в примечаниях желательно написать в чем конкретно возникали трудности)" + Environment.NewLine;
                }
                else
                {
                    para3.Range.Text = "* Заполняется воспитателями (выполняет с легкостью/ справляется/ возникают незначительные затруднения/ возникли значительные трудности/ отказ от выполнения задания/  не выполнили задания; в примечаниях желательно написать в чем конкретно возникали трудности)" + Environment.NewLine;
                }
                para3.Range.InsertParagraphAfter();
                //Save the document  

                
                string filename = Directory.GetCurrentDirectory() + "/files/" + DocumentOptions[2] + ".docx";
                if (File.Exists(filename))
                {
                    document.SaveAs2(filename);
                    document.Close(ref missing, ref missing, ref missing);
                    document = null;
                    winword.Quit(ref missing, ref missing, ref missing);
                    winword = null;
                    MessageBox.Show("Document created successfully !");
                }
                else
                {
                    filename = Directory.GetCurrentDirectory() + "/files/" + DocumentOptions[2] + DocumentOptions[1] + ".docx";
                    document.SaveAs2(filename);
                    document.Close(ref missing, ref missing, ref missing);
                    document = null;
                    winword.Quit(ref missing, ref missing, ref missing);
                    winword = null;
                    MessageBox.Show("Document created successfully !");
                }
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

        private void MainMenu_Click(object sender, RoutedEventArgs e)
        {
            MainMenuFunc();
        }

        private void MainMenuFunc()
        {
            List_Direction_UI.Visibility = Visibility.Visible;
            List_Task_UI.Visibility = Visibility.Hidden;
            List_Game_UI.Visibility = Visibility.Hidden;
            Butt_WorkDirection.Visibility = Visibility.Visible;
            Butt_Task.Visibility = Visibility.Hidden;
            Butt_Games.Visibility = Visibility.Hidden;
        }

        private void Butt_Games_Click(object sender, RoutedEventArgs e)
        {
            WorkDirection workDirection = (WorkDirection)List_Direction_UI.SelectedItem;
            Task task = (Task)List_Task_UI.SelectedItem;
            Game game = (Game)List_Game_UI.SelectedItem;
            bool error = false;
            if (List_Game_UI.SelectedItem != null)
            {
                try
                {
                    int i = 0;
                    while (i < 10)
                    {
                        if (Task_Massiv[i, workDirection.Id - 1] == null)
                        {
                            Task_Massiv[i, workDirection.Id - 1] = i + 1 + ") " + task.TaskName;
                            i = 10;
                        }
                        i++;
                    }
                }
                catch
                {
                    MessageBox.Show("Можго добавить только 10 игр в 1но направление");
                    error = true;
                }

                if (!error)
                {
                    try
                    {
                        int i = 0;
                        while (i < 10)
                        {
                            if (Game_Massiv[i, workDirection.Id - 1] == null)
                            {
                                if (game.Toys == null)
                                {
                                    Game_Massiv[i, workDirection.Id - 1] = i + 1 + ") " + game.Name + "\n описание:" + game.Description;
                                }
                                else
                                {
                                    Game_Massiv[i, workDirection.Id - 1] = i + 1 + ") " + game.Name + "\n описание:" + game.Description + "\n игрушки:" + game.Toys;
                                }
                                MessageBox.Show("Игра добавлена");
                                i = 10;
                            }
                            i++;
                        }
                    }
                    catch { }
                }
            }
            else
            {
                MessageBox.Show("Выбирете игру");
            }
            MainMenuFunc();
        }

        private void Create_Click(object sender, RoutedEventArgs e)
        {
            Document_Options document_Options = new Document_Options();
            document_Options.Owner = this;
            document_Options.ShowDialog();
        }

        public void CreateDocumentWithOptions()
        {
            CreateDocument();
        }
    }



}