﻿//
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
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml.Linq;
    using NinoDrive.Spreadsheets;
    using Libgame;

    public class WorksheetToXml
    {
        private readonly Worksheet worksheet;
        private readonly Dictionary<string, int> textColumnMap;

        public WorksheetToXml(Worksheet worksheet)
        {
            this.worksheet = worksheet;
            this.textColumnMap = new Dictionary<string, int>();
        }

        public int HeadRow { get; set; } = 1;
        public string HeadNodeType { get; set; } = "Node Type";
        public IList<string> TextColumns { get; } = new List<string> {
            "Edited Text", "Translated Text", "Original Text" };

        public XDocument Convert()
        {
            // Find the columns with the translated text.
            FindTextColumns();

            // Create the XML document.
            var xml = new XDocument();
            xml.Declaration = new XDeclaration("1.0", "utf-8", "true");

            // Create the root element from the content of the second cell.
            var root = new XElement(worksheet[0, 1].Substring("Rootname: ".Length));
            xml.Add(root);

            // Parse each row.
            ParseRow(root, HeadRow + 1, 0);
            return xml;
        }

        private void FindTextColumns()
        {
            // Find the text columns previously defined.
            textColumnMap.Clear();
            for (int i = 0; i < worksheet.Columns; i++) {
                if (TextColumns.Contains(worksheet[HeadRow, i]))
                    textColumnMap.Add(worksheet[HeadRow, i], i);
            }
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
            string[] nodeData = worksheet[row, column].Split(' ');
            var node = new XElement(nodeData[0]);

            // Process attributes.
            foreach (var attr in nodeData.Skip(1)) {
                string[] attrData = attr.Split('=');
                node.SetAttributeValue(
                    attrData[0],
                    attrData[1].Substring(1, attrData[1].Length - 2)); // Remove quotes
            }

            return node;
        }

        private string GetElementText(int row, int column)
        {
            // Get the text by looking some columns by priority.
            string text = TextColumns
                .Select(ct => !textColumnMap.ContainsKey(ct) ? null :
                            worksheet[row, textColumnMap[ct]])
                .FirstOrDefault(t => !string.IsNullOrEmpty(t));
            return text?.ToXmlString(column + 1, '[', ']');
        }
    }
}