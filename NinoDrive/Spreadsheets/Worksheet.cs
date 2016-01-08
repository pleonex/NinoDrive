//
//  Worksheet.cs
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
    using System;
    using GoogleWorksheetEntry = Google.GData.Spreadsheets.WorksheetEntry;
    using GoogleCellEntry      = Google.GData.Spreadsheets.CellEntry;

    public class Worksheet
    {
        private readonly GoogleWorksheetEntry entry;
        private string[,] cells;

        internal Worksheet(GoogleWorksheetEntry entry)
        {
            this.entry = entry;
            GetCells();
        }

        public string Name {
            get { return entry.Title.Text; }
        }

        private void GetCells()
        {
            var feed = entry.QueryCellFeed();
            cells = new string[feed.RowCount.Count, feed.ColCount.Count];
            foreach (GoogleCellEntry cell in feed.Entries)
                cells[cell.Row - 1, cell.Column - 1] = cell.Value;
        }

        public string this[int row, int col] {
            get {
                if (row < 0 || row >= cells.GetLength(0))
                    throw new ArgumentOutOfRangeException("row");

                if (col < 0 || col >= cells.GetLength(1))
                    throw new ArgumentOutOfRangeException("col");

                return cells[row, col];
            }
        }
    }
}

