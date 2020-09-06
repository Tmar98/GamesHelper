using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GamesHelper
{
    public class Task
    {
        public int Id { get; set; }
        public int Relations { get; set; }
        public string TaskName { get; set; }

        public Task()
        { }

        public Task(int id, int relations, string taskName)
        {
            Id = id; Relations = relations; TaskName = taskName;
        }
    }

    public class List_Tasks : List<Task>
    {

        public List_Tasks()
        { }

        public List_Tasks(string fileName, int id_direction)
        {
            if (File.Exists(fileName))
            {
                StreamReader sr = new StreamReader(fileName);
                string s0;
                while ((s0 = sr.ReadLine()) != null)
                {
                    string[] ar = s0.Split('*');
                    if (Convert.ToInt32(ar[1]) == id_direction)
                    {
                        Task task = new Task(Convert.ToInt32(ar[0]), Convert.ToInt32(ar[1]), ar[2]);
                        this.Add(task);
                    }
                }
                sr.Close();
                sr.Dispose();
            }
        }



    }
}