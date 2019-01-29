using Amdocs.Ginger.Plugin.Core;
using StandAloneActions;
using System;

namespace GingerOfficePluginConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting Office Plugin");

            using (GingerNodeStarter gingerNodeStarter = new GingerNodeStarter())
            {
                if (args.Length > 0)
                {
                    gingerNodeStarter.StartFromConfigFile(args[0]);  // file name 
                }
                else
                {                    
                    gingerNodeStarter.StartNode("Excel Service", new ExcelService());                 
                }
                gingerNodeStarter.Listen();
            }
        }
    }
}
