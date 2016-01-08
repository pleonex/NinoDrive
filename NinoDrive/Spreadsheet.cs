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
namespace NinoDrive
{
    using GoogleSpreadsheetEntry = Google.GData.Spreadsheets.SpreadsheetEntry;

    public class Spreadsheet
    {
        private readonly SpreadsheetsService service;
        private readonly GoogleSpreadsheetEntry entry;

        internal Spreadsheet(SpreadsheetsService service, GoogleSpreadsheetEntry entry)
        {
            this.service = service;
            this.entry = entry;
        }

        public string Title {
            get { return entry.Title.Text; }
        }
    }
}

