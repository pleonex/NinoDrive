// Copyright (C) 2015 zkarts
// Copyright (C) 2016 pleonex
namespace NinoDrive
{
    using System;
    using System.Linq;
    using NinoDrive.Spreadsheets;

    public static class Program
    {
        public static void Main()
        {
            // Get the services and start working.
            var service = new SpreadsheetsService();

            // Ask for file names.
            while (true) {
                // Get the requested SpreadhSheet title.
                Console.WriteLine();
                Console.Write("Type the file title (\"exit\" to quit)> ");
                string title = Console.ReadLine();
                if (title == "exit")
                    break;

                // Search for it
                var spreadsheets = service.SearchSpreadsheets(title).ToArray();
                var targetSheet = SelectSpreadsheet(spreadsheets);
                if (targetSheet == null)
                    continue;

                // TODO: Convert.
            }
        }

        private static Spreadsheet SelectSpreadsheet(Spreadsheet[] spreadsheets)
        {
            // Present the result and if it's not an exact match confirm.
            Console.WriteLine("{0} spreadsheets found:", spreadsheets.Length);
            for (int i = 0; i < spreadsheets.Length; i++)
                Console.WriteLine("\t{0}. {1}", i, spreadsheets[i].Title);

            // If nothing found, continue
            if (spreadsheets.Length == 0)
                return null;

            // If just a result, returns it
            if (spreadsheets.Length == 1)
                return spreadsheets[0];

            // If it's not an exact match, ask if we should continue.
            Console.Write("The title doesn't match exactly. Continue? [Y/n] ");
            string reply = Console.ReadLine();
            if (reply.ToLower() == "n")
                return null;

            // And ask for the index.
            Console.Write("Type the number of the spreadsheet in the list to select: ");
            int idx = Convert.ToInt32(Console.ReadLine());
            return (idx < 0 || idx >= spreadsheets.Length) ? null : spreadsheets[idx];
        }
    }
}