//using System;
//using System.Collections.Generic;
//using System.Windows.Forms;
//using Google.GData.Client;
//using Google.GData.Documents;
//using Google.GData.Spreadsheets;

//class Program {
//  private const string ClientId = "146394604134.apps.googleusercontent.com";
//  private const string ClientSecret = "Ivg6Pcb1ik1XVAIau_JGPjlq";
//  private const string Scope = "https://spreadsheets.google.com/feeds https://docs.google.com/feeds";
//  private const string RedirectUri = "urn:ietf:wg:oauth:2.0:oob";

//  /**
//   * A basic struct to store cell row/column information and the associated RnCn
//   * identifier.
//   */
//  private class CellAddress {
//    public uint Row;
//    public uint Col;
//    public string IdString;

//    /**
//     * Constructs a CellAddress representing the specified {@code row} and
//     * {@code col}. The IdString will be set in 'RnCn' notation.
//     */
//    public CellAddress(uint row, uint col) {
//      this.Row = row;
//      this.Col = col;
//      this.IdString = string.Format("R{0}C{1}", row, col);
//    }
//  }

//  [System.STAThreadAttribute()] //Needed to write the link to the clipboard
//  static void Main(string[] args) {
//    string[][] texts = { new string[] { "Peter", "Potter", "\t", "I will go the distance" }, new string[] { "Holly", "Holy", "\t", "I'll be there someday" } };
//    List<string[]> ssEntry = new List<string[]>(texts);

//    OAuth2Parameters oauthParams = new OAuth2Parameters();
//    oauthParams.ClientId = ClientId;
//    oauthParams.ClientSecret = ClientSecret;
//    oauthParams.RedirectUri = RedirectUri;
//    oauthParams.Scope = Scope;

//    string authorizationUrl = OAuthUtil.CreateOAuth2AuthorizationUrl(oauthParams);
//    Clipboard.SetText(authorizationUrl);
//    Console.WriteLine("A link has been copied to your clipboard. Please paste that"
//        + " link in a browser and authorize yourself on the webpage.  Once that is"
//        + " complete, type your access code and press enter to continue...");
//    oauthParams.AccessCode = Console.ReadLine();

//    OAuthUtil.GetAccessToken(oauthParams);
//    string accessToken = oauthParams.AccessToken;
//    Console.WriteLine("OAuth Access Token: " + accessToken);

//    GOAuth2RequestFactory requestFactory =
//        new GOAuth2RequestFactory(null, "NinoDrive-v1", oauthParams);

//    //Creating a spredsheetsservice and setting its requestfactory
//    SpreadsheetsService ssService = new SpreadsheetsService("NinoDrive-v1");
//    ssService.RequestFactory = requestFactory;

//    Google.GData.Spreadsheets.SpreadsheetQuery query = new Google.GData.Spreadsheets.SpreadsheetQuery();
//    SpreadsheetFeed feed = ssService.Query(query);

//    if (feed.Entries.Count == 0) {
//      DocumentEntry entry = new DocumentEntry();
//      entry.Title.Text = "Tesuto";
//      entry.Categories.Add(DocumentEntry.SPREADSHEET_CATEGORY);
//      DocumentsService docService = new DocumentsService("NinoDrive-v1");
//      docService.RequestFactory = requestFactory;
//      DocumentEntry newEntry = docService.Insert(
//          DocumentsListQuery.documentsBaseUri, entry);
//    }

//    ssService.RequestFactory = requestFactory;
//    query = new Google.GData.Spreadsheets.SpreadsheetQuery();
//    // Refresh the list of spreadsheets.
//    feed = ssService.Query(query);

//    SpreadsheetEntry spreadsheet;
//    int ssID;

//    for (ssID = 0; ssID < feed.Entries.Count; ssID++) {
//      if (feed.Entries[ssID].Title.Text == "Tesuto")
//        break;
//    }

//    if (ssID < feed.Entries.Count) {
//      spreadsheet = (SpreadsheetEntry)feed.Entries[ssID];
//    }
//    else {
//      DocumentEntry entry = new DocumentEntry();
//      entry.Title.Text = "Tesuto";
//      entry.Categories.Add(DocumentEntry.SPREADSHEET_CATEGORY);
//      DocumentsService docService = new DocumentsService("NinoDrive-v1");
//      docService.RequestFactory = requestFactory;
//      DocumentEntry newEntry = docService.Insert(
//          DocumentsListQuery.documentsBaseUri, entry);

//      ssService.RequestFactory = requestFactory;
//      query = new Google.GData.Spreadsheets.SpreadsheetQuery();
//      // Refresh the list of spreadsheets.
//      feed = ssService.Query(query);

//      for (ssID = 0; ssID < feed.Entries.Count; ssID++) {
//        if (feed.Entries[ssID].Title.Text == "Tesuto")
//          break;
//      }
//      spreadsheet = (SpreadsheetEntry)feed.Entries[ssID];
//    }

//    WorksheetFeed wsFeed = spreadsheet.Worksheets;
//    WorksheetEntry worksheet = (WorksheetEntry)wsFeed.Entries[0];

//    // Fetch the cell feed of the worksheet.
//    CellQuery cellQuery = new CellQuery(worksheet.CellFeedLink);
//    CellFeed cellFeed = ssService.Query(cellQuery);

//    // Build list of cell addresses to be filled in
//    List<CellAddress> cellAddrs = new List<CellAddress>();
//    for (uint row = 1; row <= ssEntry.Count; ++row) {
//      for (uint col = 1; col <= ssEntry[1].Length; ++col) {
//        cellAddrs.Add(new CellAddress(row, col));
//      }
//    }

//    // Prepare the update
//    // GetCellEntryMap is what makes the update fast.
//    Dictionary<String, CellEntry> cellEntries = GetCellEntryMap(ssService, cellFeed, cellAddrs);

//    CellFeed batchRequest = new CellFeed(cellQuery.Uri, ssService);
//    foreach (CellAddress cellAddr in cellAddrs) {
//      CellEntry batchEntry = cellEntries[cellAddr.IdString];

//      batchEntry.InputValue = ssEntry[(int)cellAddr.Row - 1][(int)cellAddr.Col - 1]; //ENTERED VALUE HERE!!

//      batchEntry.BatchData = new GDataBatchEntryData(cellAddr.IdString, GDataBatchOperationType.update);
//      batchRequest.Entries.Add(batchEntry);
//    }

//    // Submit the update
//    CellFeed batchResponse = (CellFeed)ssService.Batch(batchRequest, new Uri(cellFeed.Batch));

//    // Check the results
//    bool isSuccess = true;
//    foreach (CellEntry entry in batchResponse.Entries) {
//      string batchId = entry.BatchData.Id;
//      if (entry.BatchData.Status.Code != 200) {
//        isSuccess = false;
//        GDataBatchStatus status = entry.BatchData.Status;
//        Console.WriteLine("{0} failed ({1})", batchId, status.Reason);
//      }
//    }
//    Console.WriteLine(isSuccess ? "Batch operations successful." : "Batch operations failed");
//    Console.ReadLine();



//    ////Creating a new spreadsheet
//    //DocumentEntry entry = new DocumentEntry();
//    //entry.Title.Text = "Legal Contract";
//    //entry.Categories.Add(DocumentEntry.SPREADSHEET_CATEGORY);
//    //DocumentEntry newEntry = docService.Insert(
//    //    DocumentsListQuery.documentsBaseUri, entry);
//  }



//  /**
//   * Connects to the specified {@link SpreadsheetsService} and uses a batch
//   * request to retrieve a {@link CellEntry} for each cell enumerated in {@code
//   * cellAddrs}. Each cell entry is placed into a map keyed by its RnCn
//   * identifier.
//   *
//   * @param service the spreadsheet service to use.
//   * @param cellFeed the cell feed to use.
//   * @param cellAddrs list of cell addresses to be retrieved.
//   * @return a dictionary consisting of one {@link CellEntry} for each address in {@code
//   *         cellAddrs}
//   */
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
//    foreach (CellEntry entry in queryBatchResponse.Entries) {
//      cellEntryMap.Add(entry.BatchData.Id, entry);
//      Console.WriteLine("batch {0} (CellEntry: id={1} editLink={2} inputValue={3})",
//          entry.BatchData.Id, entry.Id, entry.EditUri,
//          entry.InputValue);
//    }

//    return cellEntryMap;
//  }
//}