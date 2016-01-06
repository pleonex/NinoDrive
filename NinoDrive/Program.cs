// Copyright (C) 2015 zkarts
// Copyright (C) 2016 pleonex
using System;

namespace NinoDrive
{
    public static class Program
    {
        public static void Main()
        {
            // Get the authorization
            var authorization = AuthorizationManager.Instance;

            // Get the services and start working.
            GoogleToXML go = new GoogleToXML();

            // Ask for file names.
            bool askStop = false;
            while (!askStop) {
                Console.Write("Type the file title (\"exit\" to quit)> ");

                string title = Console.ReadLine();
                askStop = (title == "exit");

                if (!askStop)
                    go.FromSpreadsheet(title);
            }
        }
    }
}