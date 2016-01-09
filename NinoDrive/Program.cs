//
//  AuthorizationManager.cs
//
//  Authors:
//       Benito Palacios Sánchez (aka pleonex) <benito356@gmail.com>
//       zkarts
//
//  The MIT License (MIT)
//
//  Copyright (c) 2016 Zkarts, Pleonex
//
//  Permission is hereby granted, free of charge, to any person obtaining a copy
//  of this software and associated documentation files (the "Software"), to deal
//  in the Software without restriction, including without limitation the rights
//  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
//  copies of the Software, and to permit persons to whom the Software is
//  furnished to do so, subject to the following conditions:
//
//  The above copyright notice and this permission notice shall be included in all
//  copies or substantial portions of the Software.
//
//  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
//  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
//  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
//  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
//  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
//  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
//  SOFTWARE.
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