//
//  Spreadsheet.cs
//
//  Author:
//       Benito Palacios Sánchez (aka pleonex) <benito356@gmail.com>
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
namespace NinoDrive.Spreadsheets
{
    using System;
    using System.Collections.Generic;
    using GoogleSpreadsheetEntry = Google.GData.Spreadsheets.SpreadsheetEntry;
    using GoogleWorksheetEntry   = Google.GData.Spreadsheets.WorksheetEntry;

    public class Spreadsheet
    {
        private readonly GoogleSpreadsheetEntry entry;
        private readonly Dictionary<int, Worksheet> worksheets;

        internal Spreadsheet(GoogleSpreadsheetEntry entry)
        {
            // Can't use an array since we don't want to call
            // entry.Worksheets.Entries.Count since it does a query veery sloooow.
            // The workaround it's to use a dictionary and get the worksheet when need.
            this.entry = entry;
            this.worksheets = new Dictionary<int, Worksheet>();
        }

        public string Title {
            get { return entry.Title.Text; }
        }

        public string Key {
            get {
                int idx = entry.Id.Uri.Content.LastIndexOf("/") + 1;
                return entry.Id.Uri.Content.Substring(idx);
            }
        }

        public DateTime Updated {
            get { return entry.Updated; }
        }

        public Worksheet this[int i] {
            get {
                if (!worksheets.ContainsKey(i)) {
                    var workEntry = entry.Worksheets.Entries[i] as GoogleWorksheetEntry;
                    worksheets[i] = new Worksheet(workEntry);
                }
                
                return worksheets[i];
            }
        }
    }
}

