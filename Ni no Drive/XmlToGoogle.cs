//using System;
//using System.Collections.Generic;
//using System.IO;
//using System.Windows.Forms;
//using System.Xml;
//using System.Xml.Linq;
//using Google.GData.Client;
//using Google.GData.Client.ResumableUpload;
//using Google.GData.Documents;
//using Google.GData.Spreadsheets;
//using Google.Spreadsheets;

//class XmlToGoogle {
//  private const string ClientId = "146394604134.apps.googleusercontent.com";
//  private const string ClientSecret = "Ivg6Pcb1ik1XVAIau_JGPjlq";
//  private const string Scope = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds";
//  private const string RedirectUri = "urn:ietf:wg:oauth:2.0:oob";

//  int tabs = 0, blocks = 3, maxDepth = 0, depth = 0, counter = 0;
//  string path, converted = "";
//  string[] nodesText;
//  List<string[]> ssEntry = new List<string[]>(); //List of lines in arrays split at \t

//  GOAuth2RequestFactory requestFactory;

//  private class CellAddress {
//    public uint Row;
//    public uint Col;
//    public string IdString;

//    public CellAddress(uint row, uint col) {
//      this.Row = row;
//      this.Col = col;
//      this.IdString = string.Format("R{0}C{1}", row, col);
//    }
//  }

//  [System.STAThreadAttribute()] //Needed to write the link to the clipboard
//  public static void Main(string[] args) {
//    if (args.Length != 0) //This is to account for several files
//    //if (true)
//        {
//      XmlToGoogle go = new XmlToGoogle();

//      //Authorization
//      OAuth2Parameters oauthParams = new OAuth2Parameters();
//      oauthParams.ClientId = ClientId;
//      oauthParams.ClientSecret = ClientSecret;
//      oauthParams.RedirectUri = RedirectUri;
//      oauthParams.Scope = Scope;

//      string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(oauthParams);
//      Clipboard.SetText(authorizationUrl);
//      Console.WriteLine("A link has been copied to your clipboard. Please paste that"
//          + " link in a browser and authorize yourself on the webpage.  Once that is"
//          + " complete, type your access code and press enter to continue...");
//      oauthParams.AccessCode = Console.ReadLine();

//      OAuthUtil.GetAccessToken(oauthParams);
//      string accessToken = oauthParams.AccessToken;
//      Console.WriteLine("OAuth Access Token: " + accessToken);

//      //Authorizing the requestfactory
//      go.requestFactory =
//          new GOAuth2RequestFactory(null, "NinoDrive-v1", oauthParams);

//      Console.WriteLine("-----------------------------------------");

//      for (int number = 0; number < args.Length; number++) {
//        Console.WriteLine("\nWorking on {0}...", Path.GetFileName(args[number]));
//        go.path = args[number];
//        go.ConvertXml(go.path); //fill the list of string arrays 'ssEntry'
//        go.ToSpreadsheet();  //use 'ssEntry' to fill the spreadsheet
//        go.Reset();
//        Console.WriteLine("-----------------------------------------");
//      }

//      Console.WriteLine("All done! Press enter to continue...");
//      Console.ReadLine();

//      //go.path = "SubQuest0004.sq.xml";
//      //go.ConvertXml(go.path); //fill the list of string arrays 'ssEntry'
//      //go.ToSpreadsheet();  //use 'ssEntry' to fill the spreadsheet
//      //go.Reset();
//    }
//    else {
//      Console.WriteLine("Drag one or several xml's onto this executable to convert them and upload them to your Google Drive.");
//      Console.ReadLine();
//    }
//  }

//  public void ToSpreadsheet() {
//    string title = Path.GetFileNameWithoutExtension(path);

//    //Create a spredsheetsservice and set its requestfactory
//    SpreadsheetsService ssService = new SpreadsheetsService("NinoDrive-v1");
//    ssService.RequestFactory = requestFactory;

//    Google.GData.Spreadsheets.SpreadsheetQuery query = new Google.GData.Spreadsheets.SpreadsheetQuery();

//    // Make a request to the API and get all spreadsheets.
//    SpreadsheetFeed feed = ssService.Query(query);

//    if (feed.Entries.Count == 0) {
//      DocumentEntry entry = new DocumentEntry();
//      entry.Title.Text = title;
//      entry.Categories.Add(DocumentEntry.SPREADSHEET_CATEGORY);
//      DocumentsService docService = new DocumentsService("NinoDrive-v1");
//      docService.RequestFactory = requestFactory;
//      DocumentEntry newEntry = docService.Insert(
//          DocumentsListQuery.documentsBaseUri, entry);
//    }

//    ssService.RequestFactory = requestFactory;
//    // Refresh the list of spreadsheets.
//    feed = ssService.Query(query);

//    SpreadsheetEntry spreadsheet;
//    int ssID;

//    for (ssID = 0; ssID < feed.Entries.Count; ssID++) {
//      if (feed.Entries[ssID].Title.Text == title)
//        break;
//    }

//    if (ssID < feed.Entries.Count) {
//      Console.WriteLine("Found spreadsheet: " + feed.Entries[ssID].Title.Text);
//      spreadsheet = (SpreadsheetEntry)feed.Entries[ssID];
//    }
//    else {
//      Console.WriteLine("Creating a new spreadsheet...");
//      DocumentEntry entry = new DocumentEntry();
//      entry.Title.Text = title;
//      entry.Categories.Add(DocumentEntry.SPREADSHEET_CATEGORY);
//      DocumentsService docService = new DocumentsService("NinoDrive-v1");
//      docService.RequestFactory = requestFactory;
//      DocumentEntry newEntry = docService.Insert(
//          DocumentsListQuery.documentsBaseUri, entry);

//      ssService.RequestFactory = requestFactory;
//      // Refresh the list of spreadsheets.
//      feed = ssService.Query(query);

//      for (ssID = 0; ssID < feed.Entries.Count; ssID++) {
//        if (feed.Entries[ssID].Title.Text == title)
//          break;
//      }
//      Console.WriteLine("Found spreadsheet: " + feed.Entries[ssID].Title.Text);
//      spreadsheet = (SpreadsheetEntry)feed.Entries[ssID];
//    }

//    WorksheetFeed wsFeed = spreadsheet.Worksheets;
//    WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[0];

//    worksheet.Title.Text = "Working Copy";
//    worksheet.Cols = (uint)ssEntry[1].Length;
//    worksheet.Rows = (uint)ssEntry.Count;
//    // Send the local representation of the worksheet to the API for
//    // modification.
//    worksheet.Update();

//    // Fetch the cell feed of the worksheet.
//    CellQuery cellQuery = new CellQuery(worksheet.CellFeedLink);
//    CellFeed cellFeed = ssService.Query(cellQuery);

//    for (int i = 0; i < ssEntry.Count; i += 400) {
//      Console.WriteLine("Working on batch " + (i / 400 + 1) + "...");
//      // Build list of cell addresses to be filled in
//      List<CellAddress> cellAddrs = new List<CellAddress>();
//      for (uint row = (uint)i + 1; row <= ssEntry.Count && row <= i + 400; ++row) {
//        for (uint col = 1; col <= ssEntry[1].Length; ++col) {
//          cellAddrs.Add(new CellAddress(row, col));
//        }
//      }

//      // Prepare the update
//      // GetCellEntryMap is what makes the update fast.
//      Dictionary<String, CellEntry> cellEntries = GetCellEntryMap(ssService, cellFeed, cellAddrs);

//      CellFeed batchRequest = new CellFeed(cellQuery.Uri, ssService);

//      foreach (CellAddress cellAddr in cellAddrs) {
//        CellEntry batchEntry = cellEntries[cellAddr.IdString];

//        batchEntry.InputValue = ssEntry[(int)cellAddr.Row - 1][(int)cellAddr.Col - 1]; //ENTERED VALUE HERE!!

//        batchEntry.BatchData = new GDataBatchEntryData(cellAddr.IdString, GDataBatchOperationType.update);
//        batchRequest.Entries.Add(batchEntry);
//      }

//      // Submit the update
//      CellFeed batchResponse = (CellFeed)ssService.Batch(batchRequest, new Uri(cellFeed.Batch));

//      // Check the results
//      bool isSuccess = true;
//      foreach (CellEntry entry in batchResponse.Entries) {
//        string batchId = entry.BatchData.Id;
//        if (entry.BatchData.Status.Code != 200) {
//          isSuccess = false;
//          GDataBatchStatus status = entry.BatchData.Status;
//          Console.WriteLine("{0} failed ({1})", batchId, status.Reason);
//        }
//      }
//      Console.WriteLine(isSuccess ? "Batch operations successful." : "Batch operations failed");
//    }
//  }

//  public void ConvertXml(string path) {
//    XDocument doc = XDocument.Load(path);
//    XElement xel = doc.Root;
//    nodesText = xel.ToString().Split('\n');

//    //Set the highest depth in the xml to maxDepth
//    SetMaxDepth(xel);

//    //Create the file info and match the warning to the Original Text position
//    string header = "\tFor new lines within a cell: CTRL + Enter." + Environment.NewLine + "Maximum length: \t\t";
//    for (int num = maxDepth; num > 1; num--)
//      header = "\t" + header;
//    header = "Filename " + Path.GetFileNameWithoutExtension(path) + "\tRootname: " + xel.Name + header;

//    ssEntry.Add(header.Split('\t'));
//    header = "";

//    //Construct the header with or without Block Type according to the xml depth
//    if (maxDepth >= 2)
//      for (int num = maxDepth; num >= 2; num--) {
//        header += "Node Type\t";
//        blocks++;
//      }

//    header += "Node Type\tOriginal Text\tTranslated Text\tTranslator\tProofreader";
//    ssEntry.Add(header.Split('\t'));

//    xel = doc.Root;
//    xel = (XElement)xel.FirstNode;
//    depth++;

//    Convert(xel);
//    while (xel.NextNode != null) {
//      xel = (XElement)xel.NextNode;
//      Convert(xel, false);
//    }

//    foreach (string s in converted.Split('\v')) //EDIT: WAS /n
//      ssEntry.Add(s.Split('\t'));
//  }

//  public void Reset() {
//    tabs = 0;
//    blocks = 3;
//    maxDepth = 0;
//    depth = 0;
//    counter = 0;
//    converted = "";
//    ssEntry = new List<string[]>();
//  }

//  public string AddAttributes(XElement xel) {
//    string content = "\t";
//    XAttribute xat;
//    if (xel.HasAttributes) {
//      content = " ";
//      xat = xel.FirstAttribute;
//      content += xat.ToString();
//      while (xat.NextAttribute != null) {
//        content += " ";
//        xat = xat.NextAttribute;
//        content += xat.ToString();
//      }
//      content += "\t";
//    }
//    tabs++;
//    return content;
//  }

//  public string CheckDepth(XElement xel) {
//    string returnTabs = "";
//    int checkDepth = 0;
//    while (xel.FirstNode != null && xel.FirstNode.NodeType != XmlNodeType.Text) {
//      xel = (XElement)xel.FirstNode;
//      checkDepth++;
//    }
//    for (int num = maxDepth - (depth + checkDepth); num > 0; num--) {
//      returnTabs += "\t";
//      tabs++;
//    }
//    return returnTabs;
//  }

//  public string FixValue(XElement xel) {
//    string value = "", addTabs = "", preTabs = "";
//    string[] valueFix;

//    tabs++;
//    while (tabs <= blocks - 2) {
//      preTabs += "\t";
//      tabs++;
//    }

//    if (xel.Value != String.Empty) {
//      valueFix = xel.Value.Split('\n');
//      if (valueFix.Length > 1) {
//        for (int num = 1; num < valueFix.Length - 1; num++) {
//          if (num > 1)
//            value += "\n";
//          string fixit = valueFix[num];
//          while (fixit[0] == '\t')
//            fixit.Replace('\t', ' ');
//          string[] str = fixit.Split(' ');
//          for (int i = 0; i < str.Length; i++)
//            if (str[i] != string.Empty)
//              value += str[i];
//        }
//        for (int num = tabs; num < blocks; num++)
//          addTabs += "\t";
//        tabs = 0;
//        return preTabs + value + addTabs + "\t\t";
//      }
//      else {
//        for (int num = tabs; num < blocks; num++)
//          addTabs += "\t";
//        tabs = 0;
//        return preTabs + xel.Value + addTabs + "\t\t";
//      }
//    }
//    else {
//      for (int num = tabs; num < blocks; num++)
//        addTabs += "\t";
//      tabs = 0;
//      return preTabs + xel.Value + addTabs + "\t\t";
//    }
//  }

//  public void Convert(XElement xel, bool called = false) {
//    counter++;
//    bool found = false;
//    int offset = 0;
//    string content = "", line = "";

//    line += xel.Name + AddAttributes(xel);
//    if (!called) {
//      while (!found) {
//        offset = 0;
//        while (nodesText[counter][offset] == ' ' && nodesText[counter][offset + 1] == ' ')
//          offset += 2;
//        if (xel.Name.ToString().Length > nodesText[counter].Length)
//          counter++;
//        else
//          if (offset + 1 + xel.Name.ToString().Length < nodesText[counter].Length)
//            if (nodesText[counter].Substring(offset + 1, xel.Name.ToString().Length) != xel.Name.ToString())
//              counter++;
//            else
//              found = true;
//          else
//            counter++;
//      }

//      for (int num = 2; nodesText[counter][num] == ' ' && nodesText[counter][num + 1] == ' '; num += 2) {
//        line = "\t" + line;
//        tabs++;
//      }
//    }

//    content += line;
//    line = "";
//    if (xel.FirstNode != null && xel.FirstNode.NodeType != XmlNodeType.Text) {
//      xel = (XElement)xel.FirstNode;
//      depth++;
//      if (converted == "" || called)
//        converted += content;
//      else
//        converted += "\v" + content;
//      Convert(xel, true);

//      if (xel.NextNode != null) {
//        XElement xel2 = (XElement)xel.NextNode;
//        Convert(xel2);
//        while (xel2.NextNode != null) {
//          xel2 = (XElement)xel2.NextNode;
//          Convert(xel2, false);
//        }
//      }

//      xel = xel.Parent;
//      depth--;
//    }
//    else {
//      content += FixValue(xel);
//      if (converted == "" || called)
//        converted += content;
//      else
//        converted += "\v" + content;
//      content = "";
//    }
//  }

//  public void SetMaxDepth(XElement xel) {
//    if (xel.FirstNode != null && xel.FirstNode.NodeType != XmlNodeType.Text) {
//      xel = (XElement)xel.FirstNode;
//      depth++;
//      if (depth > maxDepth)
//        maxDepth = depth;
//      SetMaxDepth(xel);
//      xel = xel.Parent;
//      depth--;
//    }

//    if (xel.NextNode != null) {
//      xel = (XElement)xel.NextNode;
//      SetMaxDepth(xel);
//    }
//  }


//  private static Dictionary<String, CellEntry> GetCellEntryMap(
//      SpreadsheetsService service, CellFeed cellFeed, List<CellAddress> cellAddrs) {
//    CellFeed batchRequest = new CellFeed(new Uri(cellFeed.Self), service);
//    foreach (CellAddress cellId in cellAddrs) {
//      CellEntry batchEntry = new CellEntry(cellId.Row, cellId.Col, cellId.IdString);
//      batchEntry.Id = new AtomId(string.Format("{0}/{1}", cellFeed.Self, cellId.IdString));
//      batchEntry.BatchData = new GDataBatchEntryData(cellId.IdString, GDataBatchOperationType.query);
//      batchRequest.Entries.Add(batchEntry);
//    }

//    CellFeed queryBatchResponse = (CellFeed)service.Batch(batchRequest, new Uri(cellFeed.Batch));

//    Dictionary<String, CellEntry> cellEntryMap = new Dictionary<String, CellEntry>();
//    Console.WriteLine("Writing...");
//    foreach (CellEntry entry in queryBatchResponse.Entries)
//      cellEntryMap.Add(entry.BatchData.Id, entry);

//    return cellEntryMap;
//  }
//}