//
//  SpreadsheetService.cs
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
    using System.Linq;

    using GoogleSpreadsheetsService = Google.GData.Spreadsheets.SpreadsheetsService;
    using GoogleSpreadsheetQuery    = Google.GData.Spreadsheets.SpreadsheetQuery;
    using GoogleSpreadsheetFeed     = Google.GData.Spreadsheets.SpreadsheetFeed;
    using GoogleSpreadsheetEntry    = Google.GData.Spreadsheets.SpreadsheetEntry;
    using GoogleCellQuery           = Google.GData.Spreadsheets.CellQuery;

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

        public static int MaxResults {
            get { return 500; }
        }

        public IEnumerable<Spreadsheet> SearchSpreadsheets(string title, int offset)
        {
            return SearchSpreadsheets(new GoogleSpreadsheetQuery {
                Title = title,
                StartIndex = offset
            });
        }

        public IEnumerable<Spreadsheet> SearchOldSpreadsheets(string title, int offset,
            DateTime startDate)
        {
            return SearchSpreadsheets(new GoogleSpreadsheetQuery {
                Title = title,
                StartIndex = offset,
                StartDate = startDate,
            });
        }

        private IEnumerable<Spreadsheet> SearchSpreadsheets(GoogleSpreadsheetQuery query)
        {
            Google.GData.Client.AtomEntryCollection entries;
            try {
                entries = service.Query(query).Entries;
            } catch (Google.GData.Client.GDataRequestException ex) {
                Console.WriteLine(ex);
                return new List<Spreadsheet>();
            }

            // Give priority to Spreadsheets that matches the title.
            var exactMatch = entries.FirstOrDefault(e => e.Title.Text == query.Title);
            return (exactMatch != null) ? 
                new[] { new Spreadsheet((GoogleSpreadsheetEntry)exactMatch) } :
                entries.Select(e => new Spreadsheet((GoogleSpreadsheetEntry)e));
        }

        public Worksheet RetrieveWorksheet(string sheetKey, string worksheetId)
        {
            var cellQuery = new GoogleCellQuery(sheetKey, worksheetId, "private", "full");
            var cellFeed = service.Query(cellQuery);
            return new Worksheet(worksheetId, cellFeed);
        }
    }
}

