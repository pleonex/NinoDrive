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
        public static void UpdateConfiguration(string fileInfoPath, string editInfoPath)
        {
            // Get the service.
            var service = new SpreadsheetsService();

            // Update first the file info since we don't need the spreadsheets for this.
            UpdateFileInfo(fileInfoPath);

            // Retrieve all the spreadsheets and update the edit info.
            int start = 0;
            var spreadsheets = service.SearchSpreadsheets("", start).ToList();
            do {
                UpdateEditInfo(editInfoPath, spreadsheets);

                start += SpreadsheetsService.MaxResults;
                spreadsheets = service.SearchSpreadsheets("", start).ToList();
            } while (spreadsheets.Count > 0);
        }

        private static void UpdateFileInfo(string fileInfoPath)
        {
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
            }
            doc.Save(fileInfoPath);
        }

        private static void UpdateEditInfo(
            string editInfoPath,
            IList<Spreadsheet> spreadsheets)
        {
            // Open the configuration file info.
            var doc = XDocument.Load(editInfoPath);
            var root = doc.Root;
            foreach (var xmlInfo in root.Element("Files").Elements("File")) {
                // If it's not a Nino script ignore
                string path = xmlInfo.Element("Path").Value;
                if (!path.Contains("ROM/data/map/script/") || !path.EndsWith(".bin"))
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
    }
}

