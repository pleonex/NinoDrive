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
namespace NinoDrive.Converter
{
    using System.Xml.Linq;
    using NinoDrive.Spreadsheets;
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
        private SpreadsheetsService service;
        private CacheManager cache;

        public override string FormatName {
            get { return "GDrive.Spreadsheet.Script"; }
        }

        public override void Initialize(GameFile file, params object[] parameters)
        {
            var paramTag = parameters[0] as XElement;
            spreadsheetKey = paramTag.Element("key").Value;
            worksheetId = paramTag.Element("worksheetId").Value;

            service = new SpreadsheetsService();
            cache = CacheManager.Instance;

            base.Initialize(file, parameters);
        }

        public override void Read(DataStream strIn)
        {
            // We don't need to read nothing from the game or previous files.
        }

        public override void Write(DataStream strOut)
        {
            // We don't need to write nothing.
        }

        public override void Import(params DataStream[] strIn)
        {
            // If the cache does not contain this key, it meas we don't have to update.
            if (!cache.Contains(spreadsheetKey))
                return;

            // Get the worksheet from GDrive.
            var worksheet = service.RetrieveWorksheet(spreadsheetKey, worksheetId);

            // Convert to XML
            var importedXml = new WorksheetToXml(worksheet).Convert();

            // Save the XML, since this format will be imported first than the script.
            // The script will read this override XML. We need first to dispose it.
            // This is not a safe operation to dispose from here, but should work atm.
            strIn[0].Dispose();
            importedXml.Save(this.ImportedPaths[0]);
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

