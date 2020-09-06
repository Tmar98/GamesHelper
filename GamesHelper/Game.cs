using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Serialization;

namespace GamesHelper
{
    [Serializable]
    public class Game
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string Description { get; set; }

        public string Toys { get; set; }

        public string GameType { get; set; }


        public Game()
        { }

        public Game(int id, string name, string description, string toys, string gameType)
        {
            Id = id; Name = name; Description = description; Toys = toys; GameType = gameType;
        }

        public Game(int id, string name, string description)
        {
            Id = id; Name = name; Description = description;
        }
    }

    public class Binding_Game : Game
    {
        public int Number { get; set; }

        public Binding_Game()
        { }

        public Binding_Game(int number,Game game)
        {
            Id = game.Id; Number = number; Name = game.Name; Description = game.Description; Toys = game.Toys; GameType = game.GameType;
        }
    }

    public class List_Games : List<Game>
    {
        public List_Games()
        { }

        //public List_Games(string fileName)
        //{
        //    if (File.Exists(fileName))
        //    {
        //        StreamReader sr = new StreamReader(fileName);
        //        string s0;
        //        while ((s0 = sr.ReadLine()) != null)
        //        {
        //            string[] ar = s0.Split('*');

        //            if (ar.Length == 5)
        //            {
        //                Game game = new Game(Convert.ToInt32(ar[0]), ar[1], ar[2], ar[3], ar[4]);
        //                this.Add(game);
        //            }
        //            else if (ar.Length == 3)
        //            {
        //                if ((s0 = sr.ReadLine()) != null)
        //                {
        //                    string[] re = s0.Split('*');

        //                    if (re.Length == 3)
        //                    {
        //                        Game game = new Game(Convert.ToInt32(ar[0]), ar[1], ar[2] + re[0], re[1], re[2]);
        //                        this.Add(game);
        //                    }
        //                    //else
        //                    //    MessageBox.Show("ошибка во вложенном ифе"+ i);
        //                }
        //            }
        //            //else
        //            //    MessageBox.Show("ошибка в основном ифе"+i);

        //        }
        //        sr.Close();
        //        sr.Dispose();
        //    }
        //}

        public List_Games(string fileName)
        {
            List_Games list_Games = new List_Games();
            list_Games = null;

            XmlSerializer serializer = new XmlSerializer(typeof(List_Games));

            using (StreamReader reader = new StreamReader(fileName))
            {
                this.AddRange( (List_Games)serializer.Deserialize(reader));
            }
        }
    }

    public class List_Binding_Game : List<Binding_Game>
    {

        public List_Binding_Game()
        { }

        

    }
}
