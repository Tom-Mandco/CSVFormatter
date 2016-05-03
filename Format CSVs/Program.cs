using System;
using System.Linq;
using MandCo.CSVFormatter.Applications.Programs;
using System.Collections.Generic;
using System.Threading;
using NLog;

namespace MandCo.CSVFormatter.Applications
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Logger logger = LogManager.GetCurrentClassLogger();
            string argumentsPassed = "";
            foreach(string arg in args)
            {
                argumentsPassed += "'" + arg + "'  ";
            }

            logger.Info("CSV Formatter Starting. Arguments passed: " + argumentsPassed);

            if (args.Count() != 0)
            {
                switch (args[0])
                {
                    case "Al510":
                        try
                        {
                            if (args.Count() == 2)
                            {
                                logger.Info("al510 running ...");
                                Al510_Format.Run(args[1]);
                            }
                            else
                            {
                                logger.Error("Al510 failed to run due to incorrect number of arguments. Argument count: " + args.Count());
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Al510 Failed to run - please check logs.");
                            logger.Error("Al510 failed to run.");
                            logger.Error(ex.Message);
                            logger.Error(ex.StackTrace);
                            Thread.Sleep(2500);
                        }
                        break;
                    case "Al365":
                        try
                        {
                            if (args.Count() == 2)
                            {
                                logger.Info("Al365 running ...");
                                Al365_Format.Run(args[1]);
                            }
                            else
                            {
                                logger.Error("Al365 failed to run due to incorrect number of arguments. Argument count: " + args.Count());
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Al365 Failed to run - please check logs.");
                            logger.Error("Al365 failed to run.");
                            logger.Error(ex.Message);
                            logger.Error(ex.StackTrace);
                            Thread.Sleep(2500);
                        }
                        break;
                }
            }
            else
            {
                logger.Error("Program failed to run due no arguments passed.");
                Console.WriteLine("No Program Number Entered ...");
                Console.WriteLine("This console app currently only runs for the following programs :");
                Console.Write("Al510 & Al365");
            }
            Thread.Sleep(1000);
            logger.Info("Program ended.\n");
            
        }
    }
}
