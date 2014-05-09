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

import cmdline.util.SimpleCommandLineParser;

import com.google.gdata.data.Link;
import com.google.gdata.data.acl.AclEntry;
import com.google.gdata.data.docs.DocumentListEntry;
import com.google.gdata.data.docs.RevisionEntry;
import com.google.gdata.util.AuthenticationException;
import com.google.gdata.util.ServiceException;

import java.io.IOException;
import java.io.PrintStream;
import java.text.ParseException;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.ConsoleHandler;
import java.util.logging.Level;
import java.util.logging.Logger;

import mobilehome.info.MobileHomeInfo;
import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


/**
 * An application that serves as a sample to show how the Documents List
 * Service can be used to search your documents and upload files.
 *
 * 
 * 
 * 
 */
public class DocumentListDemo {
  private DocumentList documentList;
  private PrintStream out;
  private String Park;
  private String WorksheetName;

  private static final String APPLICATION_NAME = "JavaGDataClientSampleAppV3.0";

  /**
   * The message for displaying the usage parameters.
   */
  private static final String[] USAGE_MESSAGE = {
      "Usage: java DocumentListDemo.jar --username <user> --password <pass> --park <CT|MESA> --worksheet <Month Year>",
      ""};

  private final String[] COMMAND_HELP_CREATE = {
      "create <object_type> <name>",
      "    object_type: document, spreadsheet, folder.",
      "    name: the name for the new object"};
  private final String[] COMMAND_HELP_TRASH = {
      "trash <resource_id> [delete]",
      "    resource_id: the resource id of the object to be deleted",
      "    \"delete\": Specify to permanently delete the document instead of just trashing it."};
  
  private final String[] COMMAND_HELP_HELP = {
      "help [command]", "    Weeeeeeeeeeeeee..."};
  private final String[] COMMAND_HELP_EXIT = {
      "exit", "    Exit the program."};
  private final String[] COMMAND_HELP_ERROR = {"unknown command"};
  private final String[] COMMAND_HELP_CELLLIST = {
	      "celllist", "    list contents of cell."};
  private final Map<String, String[]> HELP_MESSAGES;
  {
    HELP_MESSAGES = new HashMap<String, String[]>();
    HELP_MESSAGES.put("create", COMMAND_HELP_CREATE);
    HELP_MESSAGES.put("trash", COMMAND_HELP_TRASH);
    HELP_MESSAGES.put("celllist", COMMAND_HELP_CELLLIST);
    HELP_MESSAGES.put("help", COMMAND_HELP_HELP);
    HELP_MESSAGES.put("exit", COMMAND_HELP_EXIT);
    HELP_MESSAGES.put("error", COMMAND_HELP_ERROR);
  }

  /**
   * Constructor
   *
   * @param outputStream Stream to print output to.
   * @throws DocumentListException
   */
  public DocumentListDemo(PrintStream outputStream, String appName, String host, String park, String worksheetname)
      throws DocumentListException {
    out = outputStream;
    documentList = new DocumentList(appName, host);
    Park = park;
    WorksheetName = worksheetname;
  }

  /**
   * Authenticates the client using ClientLogin
   *
   * @param username User's email address
   * @param password User's password
   * @throws DocumentListException
   * @throws AuthenticationException
   */
  public void login(String username, String password) throws AuthenticationException,
      DocumentListException {
    documentList.login(username, password);
  }

  /**
   * Authenticates the client using AuthSub
   *
   * @param authSubToken authsub authorization token.
   * @throws DocumentListException
   * @throws AuthenticationException
   */
  public void login(String authSubToken)
      throws AuthenticationException, DocumentListException {
    documentList.loginWithAuthSubToken(authSubToken);
  }

  /**
   * Prints out the specified document entry.
   *
   * @param doc the document entry to print.
   */
  public void printDocumentEntry(DocumentListEntry doc) {
    StringBuffer output = new StringBuffer();

    output.append(" -- " + doc.getTitle().getPlainText() + " ");
    if (!doc.getParentLinks().isEmpty()) {
      for (Link link : doc.getParentLinks()) {
        output.append("[" + link.getTitle() + "] ");
      }
    }
    output.append(doc.getResourceId());

    out.println(output);
  }

  /**
   * Prints out the specified revision entry.
   *
   * @param doc the revision entry to print.
   */
  public void printRevisionEntry(RevisionEntry entry) {
    StringBuffer output = new StringBuffer();

    output.append(" -- " + entry.getTitle().getPlainText());
    output.append(", created on " + entry.getUpdated().toUiString() + " ");
    output.append(" by " + entry.getModifyingUser().getName() + " - "
        + entry.getModifyingUser().getEmail() + "\n");
    output.append("    " + entry.getHtmlLink().getHref());

    out.println(output);
  }

  /**
   * Prints out the specified ACL entry.
   *
   * @param entry the ACL entry to print.
   */
  public void printAclEntry(AclEntry entry) {
    out.println(" -- " + entry.getScope().getValue() + ": " + entry.getRole().getValue());
  }

  /*
   * Execute the test command to list contents of a cell 
   */
  private void executeCreateInvoice(int mobilepark, String ws) 
		  throws IOException, ServiceException, RowsExceededException, BiffException, WriteException, ParseException {
	  documentList.generateLateNotices(mobilepark, ws);
  }


  /**
   * Starts up the demo and prompts for commands.
   *
   * @throws ServiceException
   * @throws IOException
 * @throws WriteException 
 * @throws BiffException 
 * @throws RowsExceededException 
 * @throws ParseException 
   */
  public void run() 
		  throws IOException, ServiceException, InterruptedException, RowsExceededException, BiffException, WriteException, ParseException {
    //printMessage(WELCOME_MESSAGE);

    //BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));

    //while (executeCommand(reader)) {
    //}
    if(Park.toLowerCase().equals("ct")){
    	executeCreateInvoice(MobileHomeInfo.CT, WorksheetName);
    }else{
    	executeCreateInvoice(MobileHomeInfo.MESA, WorksheetName);
    }
  }

  /**
   * Prints out a message.
   *
   * @param msg the message to be printed.
   */
  private static void printMessage(String[] msg) {
    for (String s : msg) {
      System.out.println(s);
    }
  }

  private static void turnOnLogging() {
    // Configure the logging mechanisms
    Logger httpLogger =
        Logger.getLogger("com.google.gdata.client.http.HttpGDataRequest");
    httpLogger.setLevel(Level.ALL);
    Logger xmlLogger = Logger.getLogger("com.google.gdata.util.XmlParser");
    xmlLogger.setLevel(Level.ALL);

    // Create a log handler which prints all log events to the console
    ConsoleHandler logHandler = new ConsoleHandler();
    logHandler.setLevel(Level.ALL);
    httpLogger.addHandler(logHandler);
    xmlLogger.addHandler(logHandler);
  }

  /**
   * Runs the demo.
   *
   * @param args the command-line arguments
   *
   * @throws DocumentListException
   * @throws ServiceException
   * @throws IOException
 * @throws WriteException 
 * @throws BiffException 
 * @throws RowsExceededException 
 * @throws ParseException 
   */
  //Working example download spreadsheet:0AvjSGAuilZeZdDUzTDQyS1NIeldNWFJLaENwLVl4OEE /tmp/myxls.pdf
  //made one line change around line 333
  public static void main(String[] args)
      throws DocumentListException, IOException, ServiceException,
      InterruptedException, RowsExceededException, BiffException, WriteException, ParseException {
    SimpleCommandLineParser parser = new SimpleCommandLineParser(args);
    String authSub = parser.getValue("authSub", "auth", "a");
    String user = parser.getValue("username", "user", "u");
    String password = parser.getValue("password", "pass", "p");
    String host = parser.getValue("host", "s");
    String park = parser.getValue("park");
    String sheetname = parser.getValue("worksheet", "ws");
    boolean help = parser.containsKey("help", "h");
    
    if (host == null) {
      host = DocumentList.DEFAULT_HOST;
    }

    if (help || (user == null || password == null || park == null || sheetname == null) && authSub == null) {
      printMessage(USAGE_MESSAGE);
      System.exit(1);
    }

    if (parser.containsKey("log", "l")) {
      turnOnLogging();
    }

    DocumentListDemo demo = new DocumentListDemo(System.out, APPLICATION_NAME,
        host, park, sheetname);

    if (password != null) {
      demo.login(user, password);
    } else {
      demo.login(authSub);
    }

    demo.run();
  }
}

