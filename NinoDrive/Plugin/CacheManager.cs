//
// CacheManager.cs
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
    using System.Collections.Generic;
    using System.Linq;
    using NinoDrive.Spreadsheets;
    using Libgame;

    public class CacheManager
    {
        private static volatile CacheManager instance;
        private static readonly object syncRoot = new object();

        private readonly SpreadsheetsService service;
        private DateTime filterDate;
        private bool filter;
        private IList<Spreadsheet> updatedSheets;

        private CacheManager()
        {
            service = new SpreadsheetsService();
            CheckUpdateFitler();
            GetUpdatedFiles();
        }

        public static CacheManager Instance {
            get {
                if (instance == null) {
                    lock (syncRoot) {
                        if (instance == null)
                            instance = new CacheManager();
                    }
                }

                return instance;
            }
        }

        private void GetUpdatedFiles()
        {
            updatedSheets = new List<Spreadsheet>(
                service.SearchOldSpreadsheets("", 0, filterDate));
        }

        private void CheckUpdateFitler()
        {
            var oldFilterDate = filterDate;
            if (Configuration.GetInstance().Extras.ContainsKey("modime_filterDate")) {
                filterDate = Configuration.GetInstance().Extras["modime_filterDate"];
                filter = true;
            } else {
                filter = false;
            }

            if (filter || oldFilterDate != filterDate)
                GetUpdatedFiles();
        }

        public bool Contains(string key)
        {
            // Check if the filter date has changed.
            CheckUpdateFitler();

            // Returns if we need to update this file by not filtering or applying it.
            return !filter || updatedSheets.Any(s => s.Key == key);
        }
    }
}

