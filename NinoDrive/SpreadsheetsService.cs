//
//  SpreadsheetService.cs
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
using System;

namespace NinoDrive
{
    using System.Collections.Generic;
    using System.Linq;

    using GoogleSpreadsheetsService = Google.GData.Spreadsheets.SpreadsheetsService;
    using GoogleSpreadsheetQuery    = Google.GData.Spreadsheets.SpreadsheetQuery;
    using GoogleSpreadsheetFeed     = Google.GData.Spreadsheets.SpreadsheetFeed;
    using GoogleSpreadsheetEntry    = Google.GData.Spreadsheets.SpreadsheetEntry;

    public class SpreadsheetsService
    {
        private readonly AuthorizationManager authorization;
        private readonly GoogleSpreadsheetsService service;

        public SpreadsheetsService()
        {
            authorization = AuthorizationManager.Instance;
            service = new GoogleSpreadsheetsService(authorization.ProgramName);
            authorization.UpdateService(service);
        }

        public IEnumerable<Spreadsheet> SearchSpreadsheets(string title)
        {
            // Create the query and process it (we can search by StartDate too).
            var query = new GoogleSpreadsheetQuery {
                Title = title,
                NumberToRetrieve = 1000 };
            var entries = service.Query(query).Entries;

            // Give priority to Spreadsheets that matches the title.
            var exactMatch = entries.FirstOrDefault(e => e.Title.Text == title);
            return (exactMatch != null) ? 
                new[] { new Spreadsheet(this, (GoogleSpreadsheetEntry)exactMatch) } :
                entries.Select(e => new Spreadsheet(this, (GoogleSpreadsheetEntry)e));
        }
    }
}

