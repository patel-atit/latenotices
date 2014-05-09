/* Copyright (c) 2008 Google Inc.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */


package gdrive.api.docs;

import com.google.gdata.client.GoogleAuthTokenFactory.UserToken;
import com.google.gdata.client.GoogleService;
import com.google.gdata.client.Query;
import com.google.gdata.client.docs.DocsService;
import com.google.gdata.data.docs.DocumentListFeed;
import com.google.gdata.data.spreadsheet.CellEntry;
import com.google.gdata.data.spreadsheet.CellFeed;
import com.google.gdata.data.spreadsheet.SpreadsheetFeed;
import com.google.gdata.data.spreadsheet.WorksheetEntry;
import com.google.gdata.data.spreadsheet.WorksheetFeed;
import com.google.gdata.util.AuthenticationException;
import com.google.gdata.util.ServiceException;

import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import mobilehome.info.MobileHomeInfo;

/**
 * An application that serves as a sample to show how the Documents List Service
 * can be used to search your documents, upload and download files, change
 * sharing permission, file documents in folders, and view revisions history.
 *
 * 
 * 
 */
public class DocumentList {
  public DocsService service;
  public GoogleService spreadsheetsService;

  public static final String DEFAULT_HOST = "docs.google.com";

  public static final String SPREADSHEETS_SERVICE_NAME = "wise";
  public static final String SPREADSHEETS_HOST = "spreadsheets.google.com";

  private final String URL_FEED = "/feeds";
  private final String URL_DOCLIST_FEED = "/private/full";

  private final String URL_DEFAULT = "/default";

  private final String URL_CATEGORY_DOCUMENT = "/-/document";
  private final String URL_CATEGORY_SPREADSHEET = "/-/spreadsheet";
  private final String URL_CATEGORY_PRESENTATION = "/-/presentation";
  private final String URL_CATEGORY_STARRED = "/-/starred";
  private final String URL_CATEGORY_TRASHED = "/-/trashed";
  private final String URL_CATEGORY_FOLDER = "/-/folder";

  private String host;

  private final Map<String, String> DOWNLOAD_DOCUMENT_FORMATS;
  {
    DOWNLOAD_DOCUMENT_FORMATS = new HashMap<String, String>();
    DOWNLOAD_DOCUMENT_FORMATS.put("doc", "doc");
    DOWNLOAD_DOCUMENT_FORMATS.put("txt", "txt");
    DOWNLOAD_DOCUMENT_FORMATS.put("odt", "odt");
    DOWNLOAD_DOCUMENT_FORMATS.put("pdf", "pdf");
    DOWNLOAD_DOCUMENT_FORMATS.put("png", "png");
    DOWNLOAD_DOCUMENT_FORMATS.put("rtf", "rtf");
    DOWNLOAD_DOCUMENT_FORMATS.put("html", "html");
    DOWNLOAD_DOCUMENT_FORMATS.put("zip", "zip");
  }

  private final Map<String, String> DOWNLOAD_PRESENTATION_FORMATS;
  {
    DOWNLOAD_PRESENTATION_FORMATS = new HashMap<String, String>();
    DOWNLOAD_PRESENTATION_FORMATS.put("pdf", "pdf");
    DOWNLOAD_PRESENTATION_FORMATS.put("png", "png");
    DOWNLOAD_PRESENTATION_FORMATS.put("ppt", "ppt");
    DOWNLOAD_PRESENTATION_FORMATS.put("swf", "swf");
    DOWNLOAD_PRESENTATION_FORMATS.put("txt", "txt");
  }

  private final Map<String, String> DOWNLOAD_SPREADSHEET_FORMATS;
  {
    DOWNLOAD_SPREADSHEET_FORMATS = new HashMap<String, String>();
    DOWNLOAD_SPREADSHEET_FORMATS.put("xls", "xls");
    DOWNLOAD_SPREADSHEET_FORMATS.put("ods", "ods");
    DOWNLOAD_SPREADSHEET_FORMATS.put("pdf", "pdf");
    DOWNLOAD_SPREADSHEET_FORMATS.put("csv", "csv");
    DOWNLOAD_SPREADSHEET_FORMATS.put("tsv", "tsv");
    DOWNLOAD_SPREADSHEET_FORMATS.put("html", "html");
  }

  /**
   * Constructor.
   *
   * @param applicationName name of the application.
   *
   * @throws DocumentListException
   */
  public DocumentList(String applicationName) throws DocumentListException {
    this(applicationName, DEFAULT_HOST);
  }

  /**
   * Constructor
   *
   * @param applicationName name of the application
   * @param host the host that contains the feeds
   *
   * @throws DocumentListException
   */
  public DocumentList(String applicationName, String host) throws DocumentListException {
    if (host == null) {
      throw new DocumentListException("null passed in required parameters");
    }

    service = new DocsService(applicationName);

    // Creating a spreadsheets service is necessary for downloading spreadsheets
    spreadsheetsService = new GoogleService(SPREADSHEETS_SERVICE_NAME, applicationName);

    this.host = host;
  }

  /**
   * Set user credentials based on a username and password.
   *
   * @param user username to log in with.
   * @param pass password for the user logging in.
   *
   * @throws AuthenticationException
   * @throws DocumentListException
   */
  public void login(String user, String pass) throws AuthenticationException,
      DocumentListException {
    if (user == null || pass == null) {
      throw new DocumentListException("null login credentials");
    }

    service.setUserCredentials(user, pass);
    spreadsheetsService.setUserCredentials(user, pass);
  }

  /**
   * Allow a user to login using an AuthSub token.
   *
   * @param token the token to be used when logging in.
   *
   * @throws AuthenticationException
   * @throws DocumentListException
   */
  public void loginWithAuthSubToken(String token) throws AuthenticationException,
      DocumentListException {
    if (token == null) {
      throw new DocumentListException("null login credentials");
    }

    service.setAuthSubToken(token);
    spreadsheetsService.setAuthSubToken(token);
  }


  /**
   * Search the documents, and return a feed of docs that match.
   *
   * @param searchParameters parameters to be used in searching criteria.
   *    accepted parameters are:
   *    "q": Typical search query
   *    "alt":
   *    "author":
   *    "updated-min": Lower bound on the last time a document' content was changed.
   *    "updated-max": Upper bound on the last time a document' content was changed.
   *    "edited-min": Lower bound on the last time a document was edited by the
   *        current user. This value corresponds to the app:edited value in the
   *        Atom entry, which represents changes to the document's content or metadata.
   *    "edited-max": Upper bound on the last time a document was edited by the
   *        current user. This value corresponds to the app:edited value in the
   *        Atom entry, which represents changes to the document's content or metadata.
   *    "title": Specifies the search terms for the title of a document.
   *        This parameter used without title-exact will only submit partial queries, not exact
   *        queries.
   *    "title-exact": Specifies whether the title query should be taken as an exact string.
   *        Meaningless without title. Possible values are true and false.
   *    "opened-min": Bounds on the last time a document was opened by the current user.
   *        Use the RFC 3339 timestamp format. For example: 2005-08-09T10:57:00-08:00
   *    "opened-max": Bounds on the last time a document was opened by the current user.
   *        Use the RFC 3339 timestamp format. For example: 2005-08-09T10:57:00-08:00
   *    "owner": Searches for documents with a specific owner.
   *        Use the email address of the owner.
   *    "writer": Searches for documents which can be written to by specific users.
   *        Use a single email address or a comma separated list of email addresses.
   *    "reader": Searches for documents which can be read by specific users.
   *        Use a single email address or a comma separated list of email addresses.
   *    "showfolders": Specifies whether the query should return folders as well as documents.
   *        Possible values are true and false.
   * @param category define the category to search. (documents, spreadsheets, presentations,
   *     starred, trashed, folders)
   *
   * @throws IOException
   * @throws MalformedURLException
   * @throws ServiceException
   * @throws DocumentListException
   */
  public DocumentListFeed search(Map<String, String> searchParameters, String category)
      throws IOException, MalformedURLException, ServiceException, DocumentListException {
    if (searchParameters == null) {
      throw new DocumentListException("searchParameters null");
    }

    URL url;

    if (category == null || category.equals("")) {
      url = buildUrl(URL_DEFAULT + URL_DOCLIST_FEED);
    } else if (category.equals("documents")) {
      url = buildUrl(URL_DEFAULT + URL_DOCLIST_FEED + URL_CATEGORY_DOCUMENT);
    } else if (category.equals("spreadsheets")) {
      url = buildUrl(URL_DEFAULT + URL_DOCLIST_FEED + URL_CATEGORY_SPREADSHEET);
    } else if (category.equals("presentations")) {
      url = buildUrl(URL_DEFAULT + URL_DOCLIST_FEED + URL_CATEGORY_PRESENTATION);
    } else if (category.equals("starred")) {
      url = buildUrl(URL_DEFAULT + URL_DOCLIST_FEED + URL_CATEGORY_STARRED);
    } else if (category.equals("trashed")) {
      url = buildUrl(URL_DEFAULT + URL_DOCLIST_FEED + URL_CATEGORY_TRASHED);
    } else if (category.equals("folders")) {
      url = buildUrl(URL_DEFAULT + URL_DOCLIST_FEED + URL_CATEGORY_FOLDER);
    } else {
      throw new DocumentListException("invaild category");
    }

    Query qry = new Query(url);

    for (String key : searchParameters.keySet()) {
      qry.setStringCustomParameter(key, searchParameters.get(key));
    }

    return service.query(qry, DocumentListFeed.class);
  }

  
  public void generateLateNotices(int mobilepark, String worksheetname) 
		  throws IOException, ServiceException, RowsExceededException, BiffException, WriteException, ParseException {
	  UserToken spreadsheetsToken = (UserToken) spreadsheetsService
	            .getAuthTokenFactory().getAuthToken();
	  service.setUserToken(spreadsheetsToken.getValue());
  
	  URL SPREADSHEET_FEED_URL = new URL(
			  "https://spreadsheets.google.com/feeds/spreadsheets/private/full");
	  
	  SpreadsheetFeed feed = service.getFeed(SPREADSHEET_FEED_URL,
		        SpreadsheetFeed.class);
	  
	  List<com.google.gdata.data.spreadsheet.SpreadsheetEntry> spreadsheets = feed.getEntries();
	  
	  com.google.gdata.data.spreadsheet.SpreadsheetEntry mySSEntry = null;
	  
	  for(com.google.gdata.data.spreadsheet.SpreadsheetEntry ssentry: spreadsheets) {
		  if(ssentry.getTitle().getPlainText().equalsIgnoreCase("Rent Roll")){
			  mySSEntry = ssentry;
			  break;
		  }
	  }
	  if(mySSEntry != null){
		  System.out.println("Found \"Rent Roll\" Spreadsheet");
	  }else{
		  System.out.println("Could not find spreadsheet");
		  System.exit(-1);
	  }
	  
	  WorksheetFeed worksheetFeed = service.getFeed(mySSEntry.getWorksheetFeedUrl(), WorksheetFeed.class);
	  List<WorksheetEntry> worksheets = worksheetFeed.getEntries();
	  WorksheetEntry worksheet = null;
	  String sheetname = worksheetname;
	  
	  for (WorksheetEntry ws : worksheets) {
		  //System.out.println(ws.getTitle().getPlainText());
		  if(ws.getTitle().getPlainText().equals(sheetname)){
			  worksheet = ws;
		  }
	  }
	  
	  if(worksheet == null){
		  System.out.println("Could not find Worksheet: " + sheetname);
		  System.exit(0);
	  }else{
		  System.out.println("Found Worksheet: " + sheetname);
	  }
	  
	  int maxlots = (mobilepark == MobileHomeInfo.CT) ? MobileHomeInfo.CT_MAX_LOTS : MobileHomeInfo.MESA_MAX_LOTS;
	  
	  List<MobileHomeInfo> mobilehomeinfo = new ArrayList<MobileHomeInfo>();
	  for(int everylot = 2; everylot < maxlots; everylot++){
		  URL cellFeedUrl = new URL(worksheet.getCellFeedUrl().toString()+ "?min-row="+everylot+"&min-col=1&max-row="+everylot+"&max-col=17");
		  //URL cellFeedUrl = new URL(worksheet.getCellFeedUrl().toString()+MobileHomeInfo.LotBalanceURL(13));
		  CellFeed cellFeed = service.getFeed(cellFeedUrl, CellFeed.class);
		  
		  float lotrent=0, mhprent=0, taxes_insurance=0;
		  float previousbalance = 0;
		  float latefee = 0;
		  float credit= 0;
		  int lotnumber=0;
		  float receivedbefore5th = 0;
		  for (CellEntry cell : cellFeed.getEntries()) {
			  if(cell.getTitle().getPlainText().contains("A")){
				  lotnumber = Integer.parseInt(cell.getPlainTextContent());
			  }
			  if(cell.getTitle().getPlainText().contains("G")){
				  lotrent = Float.parseFloat(cell.getPlainTextContent());
			  }
			  if(cell.getTitle().getPlainText().contains("H")){
				  mhprent = Float.parseFloat(cell.getPlainTextContent());
			  }
			  if(cell.getTitle().getPlainText().contains("I")){
				  taxes_insurance = Float.parseFloat(cell.getPlainTextContent());
			  }
			  if(cell.getTitle().getPlainText().contains("J")){
				  previousbalance = Float.parseFloat(cell.getPlainTextContent());
			  }
			  if(cell.getTitle().getPlainText().contains("K")){
				  latefee = Float.parseFloat(cell.getPlainTextContent());
			  }	  
			  if(cell.getTitle().getPlainText().contains("L")){
				  credit = Float.parseFloat(cell.getPlainTextContent());
			  }	  
			  if(cell.getTitle().getPlainText().contains("N")){
				  receivedbefore5th = Float.parseFloat(cell.getPlainTextContent());
			  }
		  }
		
		  mobilehomeinfo.add(new MobileHomeInfo(lotrent+mhprent+taxes_insurance, 
				  								previousbalance, 
				  								latefee,
				  								credit,
				  								lotnumber,
				  								receivedbefore5th));
	  }
	  float sum = 0; 
	  for(MobileHomeInfo mh : mobilehomeinfo){
		  sum=sum+mh.TotalDue();
		  //System.out.println(mh + "Total Due: " + mh.TotalDue());
	  }
	  System.out.println("Total Rent Due: " + sum);
	  MobileHomeInfo.GenerateLateNotices(mobilehomeinfo, mobilepark);
  }

 
 
  /**
   * Gets the suffix of the resourceId. If the resourceId is
   * "document:dh3bw3j_0f7xmjhd8", "dh3bw3j_0f7xmjhd8" will be returned.
   *
   * @param resourceId the resource id to extract the suffix from.
   *
   * @throws DocumentListException
   */
  public String getResourceIdSuffix(String resourceId) throws DocumentListException {
    if (resourceId == null) {
      throw new DocumentListException("null resourceId");
    }

    if (resourceId.indexOf("%3A") != -1) {
      return resourceId.substring(resourceId.lastIndexOf("%3A") + 3);
    } else if (resourceId.indexOf(":") != -1) {
      return resourceId.substring(resourceId.lastIndexOf(":") + 1);
    }
    throw new DocumentListException("Bad resourceId");
  }

  /**
   * Gets the prefix of the resourceId. If the resourceId is
   * "document:dh3bw3j_0f7xmjhd8", "document" will be returned.
   *
   * @param resourceId the resource id to extract the suffix from.
   *
   * @throws DocumentListException
   */
  public String getResourceIdPrefix(String resourceId) throws DocumentListException {
    if (resourceId == null) {
      throw new DocumentListException("null resourceId");
    }

    if (resourceId.indexOf("%3A") != -1) {
      return resourceId.substring(0, resourceId.indexOf("%3A"));
    } else if (resourceId.indexOf(":") != -1) {
      return resourceId.substring(0, resourceId.indexOf(":"));
    } else {
      throw new DocumentListException("Bad resourceId");
    }
  }

  /**
   * Builds a URL from a patch.
   *
   * @param path the path to add to the protocol/host
   *
   * @throws MalformedURLException
   * @throws DocumentListException
   */
  private URL buildUrl(String path) throws MalformedURLException, DocumentListException {
    if (path == null) {
      throw new DocumentListException("null path");
    }

    return buildUrl(path, null);
  }

  /**
   * Builds a URL with parameters.
   *
   * @param path the path to add to the protocol/host
   * @param parameters parameters to be added to the URL.
   *
   * @throws MalformedURLException
   * @throws DocumentListException
   */
  private URL buildUrl(String path, String[] parameters)
      throws MalformedURLException, DocumentListException {
    if (path == null) {
      throw new DocumentListException("null path");
    }

    return buildUrl(host, path, parameters);
  }

  /**
   * Builds a URL with parameters.
   *
   * @param domain the domain of the server
   * @param path the path to add to the protocol/host
   * @param parameters parameters to be added to the URL.
   *
   * @throws MalformedURLException
   * @throws DocumentListException
   */
  private URL buildUrl(String domain, String path, String[] parameters)
      throws MalformedURLException, DocumentListException {
    if (path == null) {
      throw new DocumentListException("null path");
    }

    StringBuffer url = new StringBuffer();
    url.append("https://" + domain + URL_FEED + path);

    if (parameters != null && parameters.length > 0) {
      url.append("?");
      for (int i = 0; i < parameters.length; i++) {
        url.append(parameters[i]);
        if (i != (parameters.length - 1)) {
          url.append("&");
        }
      }
    }

    return new URL(url.toString());
  }
}
