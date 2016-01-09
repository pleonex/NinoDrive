//
//  AuthorizationManager.cs
//
//  Authors:
//       Benito Palacios Sánchez (aka pleonex) <benito356@gmail.com>
//       zkarts
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
namespace NinoDrive.Converter
{
    using System.Xml.Linq;
    using NinoDrive.Spreadsheets;
    using Libgame;

    public class WorksheetToXml
    {
        private const int HeadRow = 1;
        private const string HeadNodeType = "Node Type";

        private readonly Worksheet worksheet;

        public WorksheetToXml(Worksheet worksheet)
        {
            this.worksheet = worksheet;
        }

        public XDocument Convert()
        {
            var xml = new XDocument();
            xml.Declaration = new XDeclaration("1.0", "utf-8", "true");

            var root = new XElement(worksheet[0, 1].Substring("Rootname: ".Length));
            xml.Add(root);

            FindTextColumns();

            ParseRow(root, HeadRow + 1, 0);
            return xml;
        }

        private int ParseRow(XElement parent, int row, int column)
        {
            if (IsNodeColumn(row, column)) {
                do {
                    // If it's a node column, go to the next level.
                    XElement node = CreateNode(row, column);
                    parent.Add(node);

                    row = ParseRow(node, row, column + 1);
                } while (ContainsSubNodes(row, column));
            } else {
                // Get translation text.
                parent.Value = GetElementText(row, column) ?? "";
                row++;
            }

            return row;
        }

        private void FindTextColumns()
        {
            // Find the text columns previously defined.
            //throw new NotImplementedException();
        }

        private bool IsNodeColumn(int row, int column)
        {
            return worksheet[HeadRow, column] == HeadNodeType &&
                !string.IsNullOrEmpty(worksheet[row, column]);
        }

        private bool ContainsSubNodes(int row, int column)
        {
            // If there are no more rows, finish.
            if (row >= worksheet.Rows)
                return false;

            // If it's the first column continue until reach the end.
            if (column == 0)
                return true;

            // Otherwise check if all the parent columns/cells are null. This 
            // happens when they are merged cells. If some of them contains a value,
            // it means that the parent has changed so there isn't more subnodes.
            for (int c = column - 1; c >= 0; c--)
                if (worksheet[row, c] != null)
                    return false;

            return true;
        }

        private XElement CreateNode(int row, int column)
        {
            // In the row you can find the name and all the attributes.
            //throw new NotImplementedException();
            return new XElement(worksheet[row, column].Split(' ')[0]);
        }

        private string GetElementText(int row, int column)
        {
            // Get the text by looking some columns by priority.
            //throw new NotImplementedException();
            return worksheet[row, 4].ToXmlString(column + 1, '[', ']');
        }
    }
}