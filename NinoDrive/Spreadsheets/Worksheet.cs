//
//  Worksheet.cs
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
    using GoogleWorksheetEntry = Google.GData.Spreadsheets.WorksheetEntry;
    using GoogleCellEntry      = Google.GData.Spreadsheets.CellEntry;
    using GoogleCellFeed       = Google.GData.Spreadsheets.CellFeed;

    public class Worksheet
    {
        private string[,] cells;

        internal Worksheet(GoogleWorksheetEntry entry)
        {
            Name = entry.Title.Text;
            ParseCells(entry.QueryCellFeed());
        }

        internal Worksheet(string name, GoogleCellFeed feed)
        {
            Name = name;
            ParseCells(feed);
        }

        public string Name {
            get;
            private set;
        }

        public int Rows {
            get { return cells.GetLength(0); }
        }

        public int Columns {
            get { return cells.GetLength(1); }
        }

        private void ParseCells(GoogleCellFeed feed)
        {
            cells = new string[feed.RowCount.Count, feed.ColCount.Count];
            foreach (GoogleCellEntry cell in feed.Entries)
                cells[cell.Row - 1, cell.Column - 1] = cell.Value;
        }

        public string this[int row, int col] {
            get {
                if (row < 0 || row >= Rows)
                    throw new ArgumentOutOfRangeException("row");

                if (col < 0 || col >= Columns)
                    throw new ArgumentOutOfRangeException("col");

                return cells[row, col];
            }
        }
    }
}

