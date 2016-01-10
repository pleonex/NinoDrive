//
// ModimeConfigurationGenerator.cs
//
// Author:
//       Benito Palacios Sánchez (aka pleonex) <benito356@gmail.com>
//
// Copyright (c) 2016 Benito Palacios Sánchez
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in
// all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
// THE SOFTWARE.
using System;
using NinoDrive.Spreadsheets;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Linq;

namespace NinoDrive.CLI
{
    public static class ModimeConfigurationGenerator
    {
        public static void UpdateWithSpreadsheet(string fileInfoPath, string editInfoPath,
            string keySpreadsheet, string worksheetId)
        {
            // Get the service.
            var service = new SpreadsheetsService();

            // Get all the info from spreadsheet parsing the spreadsheet with keys.
            Console.WriteLine("Downloading spreadsheet...");
            var keys = service.RetrieveWorksheet(keySpreadsheet, worksheetId);
            var spreadsheets = new List<SpreadsheetData>(keys.Rows);
            for (int r = 0; r < keys.Rows; r++) {
                // Search the last non-null column.
                int keyCol = -2;
                for (int c = 0; c < keys.Columns && keyCol == -2; c++)
                    if (string.IsNullOrEmpty(keys[r, c]))
                        keyCol = c - 1;

                if (keyCol < 0)
                    continue;

                // The previous non-null column is the key and the previous previous name.
                string name = keys[r, keyCol - 1];
                string key = keys[r, keyCol];

                // Special case for tutorial, say below.
                if (name.StartsWith("Tut0"))
                    name = "Tutorial_" + name.Substring(3, 2) + "/" + name.Substring(6, 2);

                spreadsheets.Add(new SpreadsheetData(name, key));
            }

            Console.WriteLine("Update edit info...");
            UpdateEditInfo(editInfoPath, spreadsheets);
        }

        private static void UpdateEditInfo(string editInfoPath,
            IList<SpreadsheetData> spreadsheets)
        {
            // Open the configuration file info.
            var doc = XDocument.Load(editInfoPath);
            var root = doc.Root;
            foreach (var xmlInfo in root.Element("Files").Elements("File")) {
                // Search spreadsheet with the same name
                string path = xmlInfo.Element("Path").Value;
                var spreadsheet = SearchSpreadsheet(spreadsheets, path);
                if (spreadsheet == null)
                    continue;

                // For each script, create a new file type to parse the parsed script.
                var scriptInfo = new XElement("VirtualFile");
                scriptInfo.Add(new XComment("Spreadsheet for: " + path));
                scriptInfo.Add(new XElement("InternalFilter", "false"));
                scriptInfo.Add(new XElement("Type", "GDrive.Spreadsheet"));
                scriptInfo.Add(xmlInfo.Elements("Import").Where(e => e.Value.EndsWith(".xml")));
                var parameters = new XElement("Parameters");
                parameters.Add(new XElement("Key", spreadsheet.Key));
                parameters.Add(new XElement("WorksheetId", "od6")); // 0 ID
                scriptInfo.Add(parameters);
                xmlInfo.AddBeforeSelf(scriptInfo);
            }

            doc.Save(editInfoPath);
        }

        private static SpreadsheetData SearchSpreadsheet(
            IList<SpreadsheetData> spreadsheets, string path)
        {
            // Get the name of the game file.
            // Special case for tutorials, we keep the previous folder since the files
            // has common names 00.bin, 01.bin accross folders.
            int idx = path.Contains("Tutorial_") ? path.Length - 8 : path.Length - 1;
            string name = path.Substring(path.LastIndexOf('/', idx) + 1);

            // Remove extension if any (except for .sq files...).
            if (name.Contains(".") && !name.EndsWith(".sq"))
                name = name.Substring(0, name.LastIndexOf("."));

            if (name == "ARM9" || name.StartsWith("Overlay"))
                name = name.ToLower();

            // Search spreadsheet with the same name
            var sp = spreadsheets.FirstOrDefault(s => s.Title == name);
            return sp;
        }

        private class SpreadsheetData
        {
            public SpreadsheetData(string title, string key)
            {
                Title = title;
                Key = key;
            }

            public SpreadsheetData(Spreadsheet spreadsheet)
            {
                Title = spreadsheet.Title;
                Key = spreadsheet.Key;
            }

            public string Title { get; private set; }
            public string Key { get; private set; }
        }
    }
}

