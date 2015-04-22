using Maximis.Toolkit.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Maximis.Toolkit.ConsoleApp
{
    public static class ConsoleAppHelper
    {
        public static void PerformAllActions(List<AppAction> actions, string[] args)
        {
            SetupTraceListeners();

            foreach (AppAction action in actions)
            {
                if (!RunAction(action, args))
                {
                    Trace.WriteLine(string.Empty);
                    Trace.Write("Press any key to quit...");
                    Console.ReadKey(true);
                    return;
                }
            }
        }

        public static string ReadArg(string[] args, int index)
        {
            return args.Length > index ? args[index] : null;
        }

        public static void SetupTraceListeners()
        {
            Trace.Listeners.Add(new ConsoleTraceListener());
            Trace.Listeners.Add(new LogFileTraceListener(Path.Combine(Environment.CurrentDirectory, string.Format("Trace\\{0:yyyyMMddHHmmss}.txt", DateTime.Now))));
        }

        public static void ShowAppMenu(List<AppAction> actions, string textBeforeMenu, string[] args, params ConsoleKey[] menuBreaks)
        {
            SetupTraceListeners();

            while (true)
            {
                Console.Clear();
                if (!string.IsNullOrEmpty(textBeforeMenu))
                {
                    Trace.WriteLine(textBeforeMenu);
                    Trace.WriteLine(string.Empty);
                }

                foreach (AppAction action in actions)
                {
                    if (menuBreaks != null && menuBreaks.Contains(action.MenuKey))
                    {
                        Trace.WriteLine(string.Empty);
                    }
                    Trace.WriteLine(string.Format("[{0}] {1}", action.MenuKey, action.Description));
                }

                Trace.WriteLine(string.Empty);
                Trace.Write("Press Escape to quit or type a letter:");
                ConsoleKeyInfo letterChoice = Console.ReadKey(true);
                Trace.WriteLine(letterChoice.KeyChar);

                AppAction actionToRun = actions.SingleOrDefault(q => q.MenuKey == letterChoice.Key);
                if (actionToRun == null) return;

                RunAction(actionToRun, args);

                Trace.WriteLine(string.Empty);
                Trace.Write("Press any key to continue...");
                Console.ReadKey(true);
            }
        }

        private static bool RunAction(AppAction actionToRun, string[] args)
        {
            Console.Clear();
            Trace.WriteLine(string.Format("[{0}] {1}", actionToRun.MenuKey, actionToRun.Description));
            Trace.WriteLine(string.Empty);

            //try
            //{
            actionToRun.PerformAction(args);
            return true;
            //}
            //catch (Exception ex)
            //{
            //    Trace.WriteLine(string.Empty);
            //    Trace.WriteLine(string.Format("ERROR :: {0} :: {1}", ex.GetType().Name, ex.Message));
            //    Trace.WriteLine(string.Empty);
            //    Trace.WriteLine(ex.StackTrace);
            //    return false;
            //}
        }
    }
}