// Copyright (C) 2015 zkarts
// Copyright (C) 2016 pleonex
namespace NinoDrive.Converter
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Xml;
    using System.Xml.Linq;
    using NinoDrive.Spreadsheets;

    public class WorksheetToXml
    {
        int tabs = 0;
        int blocks = 3;
        int maxDepth = 0;
        int depth = 0;
        int counter = 0;
        string converted = "";
        string[] nodesText;
        List<string[]> ssEntry = new List<string[]>();
        //List of lines in arrays split at \t

        //XmlTextWriter xmlWriter;

        public void FromSpreadsheet(Worksheet worksheet)
        {
            /*
            WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[0];
            CellQuery cellQuery = new CellQuery(worksheet.CellFeedLink);
            CellFeed cellFeed = ssService.Query(cellQuery);

            int rows = cellFeed.RowCount.IntegerValue;
            int cols = cellFeed.ColCount.IntegerValue;

            var values = (from entry in cellFeed.Entries
                                         let selectableEntry = entry as CellEntry
                                         orderby selectableEntry.Row, selectableEntry.Column ascending
                                         select selectableEntry).To2DArray(
                             rows,
                             cols);

            if (values[0, 0] == null || !values[0, 0].Contains("Filename")) {
                Console.WriteLine("ERROR: Filename corrupted. Please reset the filename in the top-left column.");
                Console.WriteLine("---------------------------------");
                return;
            }
            Console.WriteLine("Making file " + values[0, 0].Substring(9) + "...");
            xmlWriter = new XmlTextWriter(".\\files\\" + values[0, 0].Substring(9) + ".xml",
                System.Text.Encoding.UTF8);
            xmlWriter.Formatting = Formatting.Indented;
            xmlWriter.Indentation = 2;
            xmlWriter.IndentChar = ' ';

            xmlWriter.WriteStartDocument();
            xmlWriter.WriteStartElement(values[0, 1].Substring(10));
            if (values[0, 1].Substring(10) == "Subtitle")
                xmlWriter.WriteAttributeString("Name", values[0, 0].Substring(9));

            uint translatedColumn = 0;
            for (uint i = 0; i < cols; i++) {
                if (values[1, i] == "Translated Text") {
                    translatedColumn = i;
                    break;
                }
            }

            string[] split;
            int depth = 0;

            //first line to initiate loop
            Console.Write("Working on row 3...");
            for (uint j = 0; j < translatedColumn - 1; j++) {
                if (!string.IsNullOrEmpty(values[2, j])) {
                    depth++;
                    split = values[2, j].Split(' ');
                    xmlWriter.WriteStartElement(split[0]);
                    if (split.Length > 1) {
                        for (int c = 1; c < split.Length; c++) {
                            var attribute = split[c];
                            while (attribute.Split('"').Count() % 2 == 0) {
                                c++;
                                attribute += " " + split[c];
                            }
                            xmlWriter.WriteAttributeString(
                                attribute.Split('=')[0],
                                attribute.Split('=')[1].Substring(1).TrimEnd('"'));
                        }
                    }
                }
            }

            string text = values[2, translatedColumn];
            if (string.IsNullOrWhiteSpace(text))
                text = values[2, translatedColumn - 1];

            if (!string.IsNullOrEmpty(text)) {
                if (text.Contains("\n")) {
                    text = "\n" + text + "\n";
                    text = text.Replace(
                        "\n",
                        "\n" + new string(
                            ' ',
                            (depth + 1) * 2));
                    text = text.Remove(text.Length - 2);
                    xmlWriter.WriteValue(text);
                } else {
                    xmlWriter.WriteValue(text);
                }
            }

            xmlWriter.WriteEndElement();
            depth--;

            for (uint i = 3; i < rows; i++) {
                Console.Write("\rWorking on row " + (i + 1) + " of " + rows + "...");
                for (uint j = 0; j < translatedColumn - 1; j++) {
                    if (!string.IsNullOrEmpty(values[i, j])) {
                        depth++;
                        while (depth - 1 > j) {
                            xmlWriter.WriteEndElement();
                            depth--;
                        }
                        split = values[i, j].Split(' ');
                        xmlWriter.WriteStartElement(split[0]);
                        if (split.Length > 1) {
                            for (int c = 1; c < split.Length; c++) {
                                if (split[c].Split('\"').Length == 2) {
                                    xmlWriter.WriteAttributeString(
                                        split[c].Split('=')[0],
                                        (split[c].Split('=')[1] + split[c + 1]).Substring(1).TrimEnd('"'));
                                    c++;
                                } else {
                                    xmlWriter.WriteAttributeString(
                                        split[c].Split('=')[0],
                                        split[c].Split('=')[1].Substring(1).TrimEnd('"'));
                                }
                            }
                        }
                    }
                }

                text = values[i, translatedColumn];
                if (string.IsNullOrWhiteSpace(text))
                    text = values[i, translatedColumn - 1];

                if (!string.IsNullOrEmpty(text) && text.Contains("\n")) {
                    text = "\n" + text + "\n";
                    text = text.Replace(
                        "\n",
                        "\n" + new string(
                            ' ',
                            (depth + 1) * 2));
                    text = text.Remove(text.Length - 2);
                    xmlWriter.WriteValue(text);
                } else {
                    xmlWriter.WriteValue(text);
                }

                xmlWriter.WriteEndElement();
                depth--;
            }
            while (depth > 0) {
                xmlWriter.WriteEndElement();
                depth--;
            }
            xmlWriter.WriteEndDocument();
            xmlWriter.Close();

            Console.WriteLine("COMPLETE!");
            Console.WriteLine("-----------------------------------------");
            */
        }

        public void ConvertXml(string path)
        {
            XDocument doc = XDocument.Load(path);
            XElement xel = doc.Root;
            nodesText = xel.ToString().Split('\n');

            //Create the file info and match the warning to the Original Text position
            string header = "\tFor new lines within a cell: CTRL + Enter." + Environment.NewLine + "Maximum length: \t\t";
            for (int num = maxDepth; num > 1; num--)
                header = "\t" + header;
            header = "Filename " + Path.GetFileNameWithoutExtension(path) + "\tRootname: " + xel.Name + header;

            ssEntry.Add(header.Split('\t'));
            header = "";

            //Construct the header with or without Block Type according to the xml depth
            if (maxDepth >= 2)
                for (int num = maxDepth; num >= 2; num--) {
                    header += "Node Type\t";
                    blocks++;
                }

            header += "Node Type\tOriginal Text\tTranslated Text\tTranslator\tProofreader";
            ssEntry.Add(header.Split('\t'));

            xel = doc.Root;
            xel = (XElement)xel.FirstNode;
            depth++;

            Convert(xel);
            while (xel.NextNode != null) {
                xel = (XElement)xel.NextNode;
                Convert(xel, false);
            }

            foreach (string s in converted.Split('\v')) //EDIT: WAS /n
          ssEntry.Add(s.Split('\t'));
        }

        public void Reset()
        {
            tabs = 0;
            blocks = 3;
            maxDepth = 0;
            depth = 0;
            counter = 0;
            converted = "";
            ssEntry = new List<string[]>();
        }

        public string AddAttributes(XElement xel)
        {
            string content = "\t";
            XAttribute xat;
            if (xel.HasAttributes) {
                content = " ";
                xat = xel.FirstAttribute;
                content += xat.ToString();
                while (xat.NextAttribute != null) {
                    content += " ";
                    xat = xat.NextAttribute;
                    content += xat.ToString();
                }
                content += "\t";
            }
            tabs++;
            return content;
        }

        public string CheckDepth(XElement xel)
        {
            string returnTabs = "";
            int checkDepth = 0;
            while (xel.FirstNode != null && xel.FirstNode.NodeType != XmlNodeType.Text) {
                xel = (XElement)xel.FirstNode;
                checkDepth++;
            }
            for (int num = maxDepth - (depth + checkDepth); num > 0; num--) {
                returnTabs += "\t";
                tabs++;
            }
            return returnTabs;
        }

        public string FixValue(XElement xel)
        {
            string value = "", addTabs = "", preTabs = "";
            string[] valueFix;

            tabs++;
            while (tabs <= blocks - 2) {
                preTabs += "\t";
                tabs++;
            }

            if (xel.Value != String.Empty) {
                valueFix = xel.Value.Split('\n');
                if (valueFix.Length > 1) {
                    for (int num = 1; num < valueFix.Length - 1; num++) {
                        if (num > 1)
                            value += "\n";
                        string fixit = valueFix[num];
                        while (fixit[0] == '\t')
                            fixit.Replace('\t', ' ');
                        string[] str = fixit.Split(' ');
                        for (int i = 0; i < str.Length; i++)
                            if (str[i] != string.Empty)
                                value += str[i];
                    }
                    for (int num = tabs; num < blocks; num++)
                        addTabs += "\t";
                    tabs = 0;
                    return preTabs + value + addTabs + "\t\t";
                } else {
                    for (int num = tabs; num < blocks; num++)
                        addTabs += "\t";
                    tabs = 0;
                    return preTabs + xel.Value + addTabs + "\t\t";
                }
            } else {
                for (int num = tabs; num < blocks; num++)
                    addTabs += "\t";
                tabs = 0;
                return preTabs + xel.Value + addTabs + "\t\t";
            }
        }

        public void Convert(XElement xel, bool called = false)
        {
            counter++;
            bool found = false;
            int offset = 0;
            string content = "", line = "";

            line += xel.Name + AddAttributes(xel);
            if (!called) {
                while (!found) {
                    offset = 0;
                    while (nodesText[counter][offset] == ' ' && nodesText[counter][offset + 1] == ' ')
                        offset += 2;
                    if (xel.Name.ToString().Length > nodesText[counter].Length)
                        counter++;
                    else if (offset + 1 + xel.Name.ToString().Length < nodesText[counter].Length)
                    if (nodesText[counter].Substring(offset + 1, xel.Name.ToString().Length) != xel.Name.ToString())
                        counter++;
                    else
                        found = true;
                    else
                        counter++;
                }

                for (int num = 2; nodesText[counter][num] == ' ' && nodesText[counter][num + 1] == ' '; num += 2) {
                    line = "\t" + line;
                    tabs++;
                }
            }

            content += line;
            line = "";
            if (xel.FirstNode != null && xel.FirstNode.NodeType != XmlNodeType.Text) {
                xel = (XElement)xel.FirstNode;
                depth++;
                if (converted == "" || called)
                    converted += content;
                else
                    converted += "\v" + content;
                Convert(xel, true);

                if (xel.NextNode != null) {
                    XElement xel2 = (XElement)xel.NextNode;
                    Convert(xel2);
                    while (xel2.NextNode != null) {
                        xel2 = (XElement)xel2.NextNode;
                        Convert(xel2, false);
                    }
                }

                xel = xel.Parent;
                depth--;
            } else {
                content += FixValue(xel);
                if (converted == "" || called)
                    converted += content;
                else
                    converted += "\v" + content;
                content = "";
            }
        }
    }

    static class MyExtensions
    {
        public static string[,] To2DArray<T>(
            this IEnumerable<T> input,
            int rows,
            int cols) where T : Google.GData.Spreadsheets.CellEntry
        {
            string[,] ret = new string[rows, cols];
            foreach (T t in input) {
                ret[t.Row - 1, t.Column - 1] = t.Value;
            }
            return ret;
        }
    }
}