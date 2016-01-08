﻿//
//  Spreadsheet.cs
//
//  Author:
//       Benito Palacios Sánchez (aka pleonex) <benito356@gmail.com>
//
//  Copyright (c) 2016 Benito Palacios Sánchez
//
//  This program is free software: you can redistribute it and/or modify
//  it under the terms of the GNU General Public License as published by
//  the Free Software Foundation, either version 3 of the License, or
//  (at your option) any later version.
//
//  This program is distributed in the hope that it will be useful,
//  but WITHOUT ANY WARRANTY; without even the implied warranty of
//  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
//  GNU General Public License for more details.
//
//  You should have received a copy of the GNU General Public License
//  along with this program.  If not, see <http://www.gnu.org/licenses/>.
namespace NinoDrive.Spreadsheets
{
    using GoogleSpreadsheetEntry = Google.GData.Spreadsheets.SpreadsheetEntry;
    using GoogleWorksheetEntry   = Google.GData.Spreadsheets.WorksheetEntry;

    public class Spreadsheet
    {
        private readonly GoogleSpreadsheetEntry entry;
        private readonly Worksheet[] worksheets;

        internal Spreadsheet(GoogleSpreadsheetEntry entry)
        {
            this.entry = entry;
            this.worksheets = new Worksheet[entry.Worksheets.Entries.Count];
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

        public Worksheet this[int i] {
            get {
                if (worksheets[i] == null)
                    worksheets[i] =
                        new Worksheet((GoogleWorksheetEntry)entry.Worksheets.Entries[i]);
                
                return worksheets[i];
            }
        }
    }
}

