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
        public static void UpdateSearching(string fileInfoPath, string editInfoPath)
        {
            // Get the service.
            var service = new SpreadsheetsService();

            // Update first the file info since we don't need the spreadsheets for this.
            var scriptFiles = UpdateFileInfo(fileInfoPath);

            // Retrieve all the spreadsheets and update the edit info.
            int start = 0;
            var spreadsheets = service.SearchSpreadsheets("", start)
                .Select(s => new SpreadsheetData(s)).ToList();
            do {
                UpdateEditInfo(editInfoPath, spreadsheets, scriptFiles);

                start += SpreadsheetsService.MaxResults;
                spreadsheets = service.SearchSpreadsheets("", start)
                    .Select(s => new SpreadsheetData(s)).ToList();
            } while (spreadsheets.Count > 0);
        }

        public static void UpdateWithSpreadsheet(string fileInfoPath, string editInfoPath,
            string keySpreadsheet, string worksheetId)
        {
            // Get the service.
            var service = new SpreadsheetsService();

            // Update first the file info since we don't need the spreadsheets for this.
            var scriptFiles = UpdateFileInfo(fileInfoPath);

            // Get all the info from spreadsheet parsing the spreadsheet keyed.
            var keys = service.RetrieveWorksheet(keySpreadsheet, worksheetId);
            var spreadsheets = new List<SpreadsheetData>(keys.Rows);
            for (int r = 0; r < keys.Rows; r++) {
                int keyCol = -2;
                for (int c = 0; c < keys.Columns && keyCol == -2; c++)
                    if (string.IsNullOrEmpty(keys[r, c]))
                        keyCol = c - 1;

                if (keyCol < 0)
                    continue;

                spreadsheets.Add(new SpreadsheetData(keys[r, keyCol - 1], keys[r, keyCol]));
            }

            UpdateEditInfo(editInfoPath, spreadsheets, scriptFiles);
        }

        private static IList<string> UpdateFileInfo(string fileInfoPath)
        {
            var scriptFiles = new List<string>();

            // Open the configuration file info.
            var doc = XDocument.Load(fileInfoPath);
            var root = doc.Root;
            foreach (var xmlInfo in root.Element("Files").Elements("FileInfo")) {
                // If it's not a Nino script ignore
                string type = xmlInfo.Element("Type").Value;
                if (type != "Ninokuni.Script")
                    continue;

                // For each script, create a new file type to parse the parsed script.
                string path = xmlInfo.Element("Path").Value;
                var scriptInfo = new XElement("FileInfo");
                scriptInfo.Add(new XElement("Path", path + "/script"));
                scriptInfo.Add(new XElement("Type", "GDrive.Spreadsheet.Script"));
                scriptInfo.Add(new XElement("DependsOn", path));
                xmlInfo.AddAfterSelf(scriptInfo);

                scriptFiles.Add(path);
            }
            doc.Save(fileInfoPath);

            return scriptFiles;
        }

        private static void UpdateEditInfo(string editInfoPath,
            IList<SpreadsheetData> spreadsheets, IList<string> scriptFiles)
        {
            // Open the configuration file info.
            var doc = XDocument.Load(editInfoPath);
            var root = doc.Root;
            foreach (var xmlInfo in root.Element("Files").Elements("File")) {
                // If it's not a Nino script ignore
                string path = xmlInfo.Element("Path").Value;
                if (!scriptFiles.Contains(path))
                    continue;

                // Search spreadsheet with the same name
                string name = path.Substring(path.LastIndexOf('/') + 1);
                name = name.Substring(0, name.LastIndexOf("."));
                var spreadsheet = spreadsheets.FirstOrDefault(s => s.Title == name);
                if (spreadsheet == null)
                    continue;

                // For each script, create a new file type to parse the parsed script.
                var scriptInfo = new XElement("File");
                scriptInfo.Add(new XElement("Path", path + "/script"));
                scriptInfo.Add(new XElement("InternalFilter", "false"));
                scriptInfo.Add(xmlInfo.Element("Import"));
                var parameters = new XElement("Parameters");
                parameters.Add(new XElement("Key", spreadsheet.Key));
                parameters.Add(new XElement("WorksheetId", "od6")); // This is the ID of 0
                scriptInfo.Add(parameters);
                xmlInfo.AddBeforeSelf(scriptInfo);
            }
            doc.Save(editInfoPath);
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

