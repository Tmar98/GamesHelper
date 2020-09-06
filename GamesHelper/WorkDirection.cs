using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GamesHelper
{
    public class WorkDirection
    {

    public int Id { get; set; }
    public string Work { get; set; }


    public WorkDirection()
    { }

    public WorkDirection(string work)
    {
        Work = work;
    }

}

public class List_WorkDirection : List<WorkDirection>
{


    public List_WorkDirection()
    { }

    public List_WorkDirection(string fileName)
    {
        if (File.Exists(fileName))
        {
            StreamReader sr = new StreamReader(fileName);
            string s0;
            while ((s0 = sr.ReadLine()) != null)
            {
                    WorkDirection workDirection = new WorkDirection(s0);
                    workDirection.Id = this.Count + 1;
                this.Add(workDirection);
            }
            sr.Close();
            sr.Dispose();
        }
    }
}
}
