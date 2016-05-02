using System;
using System.Linq;
using MandCo.Applications.CSVFormatter.Programs;
using System.Collections.Generic;
using System.Threading;


namespace MandCo.Applications.CSVFormatter
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Count() != 0)
            {
                switch (args[0])
                {
                    case "Al510":
                        if (args.Count() == 2)
                        {
                            try
                            {
                                Al510_Format.Run(args[1]);
                            }
                            catch
                            {

                            }
                        }
                        break;
                    case "Al365":
                        if (args.Count() == 2)
                        {
                            try
                            {
                                Al365_Format.Run(args[1]);
                            }
                            catch
                            {

                            }
                        }
                        break;
                }
            }
            else
            {
                Console.WriteLine("No Program Number Entered ...");
                Console.WriteLine("This console app currently only runs for the following programs :");
                Console.Write("Al510 & Al365");
            }
            Thread.Sleep(1000);
        }
    }
}
