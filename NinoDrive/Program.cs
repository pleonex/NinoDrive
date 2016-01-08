// Copyright (C) 2015 zkarts
// Copyright (C) 2016 pleonex
using System;
using System.Linq;

namespace NinoDrive
{
    public static class Program
    {
        public static void Main()
        {
            // Get the authorization
            var authorization = AuthorizationManager.Instance;

            // Get the services and start working.
            var service = new SpreadsheetsService();

            // Ask for file names.
            while (true) {
                // Get the requested SpreadhSheet title.
                Console.Write("Type the file title (\"exit\" to quit)> ");
                string title = Console.ReadLine();
                if (title == "exit")
                    break;

                // Search for it
                var spreadsheets = service.SearchSpreadsheets(title).ToArray();

                // Present the result and if it's not an exact match confirm.
                Console.WriteLine("\t{0} spreadsheets found:", spreadsheets.Length);
                foreach (var sheet in spreadsheets)
                    Console.WriteLine("\t* {0}", sheet.Title);

                // If nothing found, continue
                if (spreadsheets.Length == 0)
                    continue;

                // If it's not an exact match, ask if we should continue.
                if (spreadsheets.Length > 1 || spreadsheets[0].Title != title) {
                    Console.Write("The title doesn't match exactly. Continue? [Y/N] ");
                    string reply = Console.ReadLine();
                    if (reply.ToLower() != "y")
                        continue;
                }
            }
        }
    }
}