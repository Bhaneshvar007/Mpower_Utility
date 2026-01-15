 
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTSUtilityPanFolioMaster
{
    public class Program
    {
        static void Main(string[] args)
        {
            FilePickingUtility filePickingUtility = new FilePickingUtility();
            Console.WriteLine("Process Started !!");
            filePickingUtility.ProcessUtilityAsyncs();
 
        }
    }
}