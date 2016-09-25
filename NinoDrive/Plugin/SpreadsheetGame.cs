//
// SpreadsheetGame.cs
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
namespace NinoDrive.Plugin
{
    using System;
    using System.IO;
    using System.Xml.Linq;
    using NinoDrive.Spreadsheets;
    using NinoDrive.Converter;
    using Libgame;
    using Libgame.IO;
    using Mono.Addins;

    /// <summary>
    /// Spreadsheet converter wrapper for Libgame.
    /// </summary>
    [Extension]
    public class SpreadsheetGame : Format
    {
        private string spreadsheetKey;
        private string worksheetId;
        private string output;
        private SpreadsheetsService service;
        private CacheManager cache;

        public override string FormatName {
            get { return "GDrive.Spreadsheet"; }
        }

        public override void Initialize(GameFile file, params object[] parameters)
        {
            var paramTag = parameters[0] as XElement;
            spreadsheetKey = paramTag.Element("Key").Value;
            worksheetId = paramTag.Element("WorksheetId").Value;
            output = Configuration.GetInstance().ResolvePath(
                paramTag.Element("Output").Value);

            service = new SpreadsheetsService();
            cache = CacheManager.Instance;

            base.Initialize(file, parameters);
        }

        public override void Read(DataStream strIn)
        {
            // We don't need to read nothing from the game or previous files.
            // IgnorePath should be set to true because this is a virtual file.
            // In this case this shouldn't be called.
            throw new NotSupportedException();
        }

        public override void Write(DataStream strOut)
        {
            // Same as read method.
            throw new NotSupportedException();
        }

        public override void Import(params DataStream[] strIn)
        {
            // If the cache does not contain this key, it meas we don't have to update.
            if (!cache.Contains(spreadsheetKey))
                return;

            // Get the worksheet from GDrive.
            Worksheet worksheet;
            try {
                worksheet = service.RetrieveWorksheet(spreadsheetKey, worksheetId);
            } catch (Google.GData.Client.GDataRequestException) {
                // Capture errors since if there isn't Internet connection
                // we don't want to abort the process.
                Console.WriteLine("ERROR: Can't get spreadhseet {0}", spreadsheetKey);
                return;
            }

            // Convert to XML
            var importedXml = new WorksheetToXml(worksheet).Convert();

            // Save to the path creating the directory if needed.
            Directory.CreateDirectory(Path.GetDirectoryName(output));
            importedXml.Save(output);
        }

        public override void Export(params DataStream[] strOut)
        {
            // Here we would export the XML to GDrive.
        }

        protected override void Dispose(bool freeManagedResourcesAlso)
        {
        }
    }
}

