import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.*;
import com.google.api.services.sheets.v4.Sheets;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.event.PopupMenuEvent;
import javax.swing.event.PopupMenuListener;
import javax.swing.table.TableRowSorter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;

/**
 * <p>A utility class for connecting to Google Drive and performing common, specific tasks for instances of
 * <code>PritzkerEHR</code>.</p>
 * @author Saieesh Rao, Pritzker School of Medicine '21
 * @version 1.3
 */
public class SheetsHandler {

    private final String APPLICATION_NAME;// Application name.
    private final java.io.File DATA_STORE_DIR; //directory to stort credentials
    private FileDataStoreFactory DATA_STORE_FACTORY; //instance of the {@link FileDataStoreFactory}.
    private final JsonFactory JSON_FACTORY; //instance of the JSON factory
    private HttpTransport HTTP_TRANSPORT; //instance of the HTTP transport

    private String spreadsheetID;
    private Credential credential;
    private List<List<Object>> values;
    /**
     * <p>A variable that contains the same information as <code>values</code>. Formerly a subset of <code>values</code>
     * containing data for interviewees whose interview date had not yet passed; however, this caused problems when
     * interviewees mis-entered their interview dates and so their requests didn't show up in Pritzker EHR.</p>
     */
    private ArrayList<ArrayList<Object>> upcomingGuests;
    private Object[][] upcomingReceiptArray;
    private Object[][] upcomingPairingArray;
    private Object[][] doneRequestArray;
    private Object[][] ignoredRequestArray;

    private static final String PATH_ID = "1fWnFRurCA23sNTy5lIc35qXSoZ165DM4tfu0RJv2COk";
    private String pleaEmailSignature;
    private String receiptEmailText;
    private String pairingEmailText;
    private String pleaEmailText;

    private final static String[] indexToLetter = new String[] {"A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};
    private final static String[] fieldNames = new String[] {"Interviewee Name","Interviewee Gender","Interview Date","Hosting Date","Interviewee Email",
            "Interviewee Phone","Undergraduate School","Preferred Host Gender","Allergies","Interest Groups"};
    private PritzkerEHR pritzkerEHR;

    /** Global instance of the scopes required by this quickstart.
     *
     * If modifying these scopes, delete your previously saved credentials
     * at ~/.credentials/sheets.googleapis.com-java-quickstart
     */
    private final List<String> SCOPES;

    private int forReceipt;
    private int forPairing;

    /**
     * <p>Constructor for instances of <code>SheetsHandler</code>. Some initialization info, particularly the
     * identity of the Google Sheet containing interviewee info, is read from the Google Sheet "PritzkerHosting_PATH".
     * </p>
     * @param pritzkerEHR the associated instance of PritzkerEHR
     */
    public SheetsHandler(PritzkerEHR pritzkerEHR){
        this.pritzkerEHR = pritzkerEHR;
        //default initialization for interviewee counter variables; -1 is an impossible value for these variables
        forReceipt = -1;    //# interviewees awaiting a receipt email
        forPairing = -1;    //# interviewees awaiting a pairing email
        APPLICATION_NAME = "Google Sheets API Java SheetsHandler"; //name of SheetsHandler application
        DATA_STORE_DIR = new java.io.File(System.getProperty("user.home"), ".credentials/pritzker_hosting_helper"); //directory to stort credentials
        JSON_FACTORY = JacksonFactory.getDefaultInstance(); //instance of the JSON factory
        SCOPES = Arrays.asList(SheetsScopes.SPREADSHEETS); //request for permission to access spreadsheets on Google Drive
        //connect to Google Drive
        try {
            HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
            DATA_STORE_FACTORY = new FileDataStoreFactory(DATA_STORE_DIR);
            credential = authorize();
        } catch (Throwable t) {
            t.printStackTrace();
            JOptionPane.showMessageDialog(pritzkerEHR,"Pritzker EHR has just had trouble connecting to Google Drive. Please wait two seconds and try again.","Oops!",JOptionPane.WARNING_MESSAGE,pritzkerEHR.getWindowIcon());
            System.exit(1);
        }
        //load presets from PritzkerHosting_PATH
        loadPATH();
        //read interviewee info from the Google Sheets
        updateValues();
        //update interviewee info in Sheet2 with the interviewee submitted info in Sheet1 that was downloaded in updateValues()
        updateSheet2();
    }

    /**
     * <p>Accessor method for the ID of the Google Sheet containing interviewee info.</p>
     * @return <p><code>String</code> representing the Google Sheet ID being used by the <code>SheetsHandler</code>.
     * </p>
     */
    public String getSpreadsheetID(){
        return spreadsheetID;
    }

    /**
     * <p>Accessor method for the array of interviewee info.</p>
     * @return a <code>List</code> of <code>List</code> containing <code>String</code> representations of spreadsheet
     * cell contents
     */
    public List<List<Object>> getValues(){
        return values;
    }

    /**
     * <p>Accessor method for the number of interviewees awaiting a receipt email.</p>
     * @return the number of interviewees awaiting a receipt email
     */
    public int numberForReceipt(){
        return forReceipt;
    }

    /**
     * <p>Accessor method for the number of interviewees awaiting a host pairing email.</p>
     * @return the number of interviewees awaiting a host pairing email.
     */
    public int numberForPairing(){
        return forPairing;
    }

    /**
     * <p>Constructs a JTable containing the info of interviewees awaiting a receipt email.</p>
     * <p>The JTable contains 4 columns for 1) hosting date, 2) interviewee name, 3) interviewee email, and
     * 4) Google Sheet row number. The last column is minimized and uneditable, and a custom
     * <code>TableRowSorter</code> is applied to the hosting date column - this is done by calling
     * <code>conformThreeColumnTable(...)</code>. The JTable is also constructed such that a customized JPopupMenu is
     * activated upon the JTable being right-clicked.</p>
     * @return the JTable for interviewees awaiting a receipt email
     */
    public JTable getReceiptJTable(){
        Object[][] upcomingArray = new Object[upcomingGuests.size()][upcomingGuests.get(0).size()];
        int ctr = 0;
        for(int i = 0; i < upcomingGuests.size(); i++){
            if(upcomingGuests.get(i).get(12).equals("0")){
                ctr++;
            }
            for(int j = 0; j < upcomingGuests.get(0).size(); j++){
                upcomingArray[i][j] = upcomingGuests.get(i).get(j);
            }
        }
        sortArrayByDate(upcomingArray,4);
        Object[][] upcomingArrayTight = new Object[ctr][4];
        upcomingReceiptArray = new Object[ctr][upcomingArray.length];
        int j = 0;
        for(int i = 0; i < upcomingArray.length; i++){
            if(upcomingArray[i][12].equals("0")) {
                upcomingArrayTight[j][0] = upcomingArray[i][4];   //hosting night
                upcomingArrayTight[j][1] = upcomingArray[i][1];   //interviewee name
                upcomingArrayTight[j][2] = upcomingArray[i][5];   //interviewee email
                upcomingArrayTight[j][3] = upcomingArray[i][13];  //spreadsheet row (for writing to sheet)
                upcomingReceiptArray[j] = upcomingArray[i];
                j++;
            }
        }
        forReceipt = j;
        String[] titlesTight = new String[] {"Hosting Date","Interviewee Name","Interviewee Email",""};
        JTable jtable = new JTable(upcomingArrayTight,titlesTight) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return column != 3;
            }
        };
        conformThreeColumnTable(jtable);
        jtable.setComponentPopupMenu(new TablePopupMenu(jtable));
        return jtable;
    }

    /**
     * <p>Constructs a JTable containing the info of interviewees awaiting a pairing email.</p>
     * <p>The JTable contains 6 columns for 1) host name, 2) host email, 3) hosting date, 4) interviewee name,
     * 5) interviewee email, and 6) Google Sheet row number. The last column is minimized and uneditable, and a custom
     * <code>TableRowSorter</code> is applied to the hosting date column - this is done by calling
     * <code>conformFiveColumnTable(...)</code>. The JTable is also constructed such that a customized JPopupMenu is
     * activated upon the JTable being right-clicked.</p>
     * @return the JTable for interviewees awaiting a pairing email
     */
    public JTable getPairingJTable(){
        Object[][] upcomingArray = new Object[upcomingGuests.size()][upcomingGuests.get(0).size()];
        int ctr = 0;
        for(int i = 0; i < upcomingGuests.size(); i++){
            if(upcomingGuests.get(i).get(upcomingGuests.get(i).size()-2).equals("1")){
                ctr++;
            }
            for(int j = 0; j < upcomingGuests.get(i).size(); j++){
                upcomingArray[i][j] = upcomingGuests.get(i).get(j);
            }
        }
        sortArrayByDate(upcomingArray,4);
        Object[][] upcomingArrayTight = new Object[ctr][6];
        upcomingPairingArray = new Object[ctr][13];
        int j = 0;
        for(int i = 0; i < upcomingArray.length; i++){
            if(upcomingArray[i][12] != null && upcomingArray[i][12].equals("1")) {
                upcomingArrayTight[j][0] = upcomingArray[i][0];   //host name
                upcomingArrayTight[j][1] = upcomingArray[i][11];  //host email
                upcomingArrayTight[j][2] = upcomingArray[i][4];   //hosting night
                upcomingArrayTight[j][3] = upcomingArray[i][1];   //interviewee name
                upcomingArrayTight[j][4] = upcomingArray[i][5];   //interviewee email
                upcomingArrayTight[j][5] = upcomingArray[i][13];  //spreadsheet row (for writing to sheet)
                upcomingPairingArray[j] = upcomingArray[i];
                j++;
            }
        }
        forPairing = j;
        String[] titlesTight = new String[] {"Host Name","Host Emails","Hosting Date","Interviewee Name","Interviewee Email",""};
        JTable jtable = new JTable(upcomingArrayTight,titlesTight) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return column != 5;
            }
        };
        conformFiveColumnTable(jtable);
        jtable.setComponentPopupMenu(new TablePopupMenu(jtable));
        return jtable;
    }

    /**
     * <p>Constructs a JTable containing the info of interviewees awaiting a pairing email.</p>
     * <p>The JTable contains 6 columns for 1) host name, 2) host email, 3) hosting date, 4) interviewee name,
     * 5) interviewee email, and 6) Google Sheet row number. The last column is minimized and uneditable, and a custom
     * <code>TableRowSorter</code> is applied to the hosting date column - this is done by calling
     * <code>conformFiveColumnTable(...)</code>. The JTable is also constructed such that a customized JPopupMenu is
     * activated upon the JTable being right-clicked.</p>
     * @return the JTable for interviewees awaiting a pairing email
     */
    public JTable getDoneJTable(){
        Object[][] upcomingArray = new Object[values.size()-1][values.get(0).size()];
        int ctr = 0;
        for(int i = 1; i < values.size(); i++){
            if(values.get(i).get(values.get(i).size()-2).equals("2")){
                ctr++;
            }
            for(int j = 0; j < values.get(i).size(); j++){
                upcomingArray[i-1][j] = values.get(i).get(j);
            }
        }
        sortArrayByDate(upcomingArray,4);
        Object[][] upcomingArrayTight = new Object[ctr][6];
        doneRequestArray = new Object[ctr][13];
        int j = 0;
        for(int i = 0; i < upcomingArray.length; i++){
            if(upcomingArray[i][12] != null && upcomingArray[i][12].equals("2")) {
                upcomingArrayTight[j][0] = upcomingArray[i][0];   //host name
                upcomingArrayTight[j][1] = upcomingArray[i][11];  //host email
                upcomingArrayTight[j][2] = upcomingArray[i][4];   //hosting night
                upcomingArrayTight[j][3] = upcomingArray[i][1];   //interviewee name
                upcomingArrayTight[j][4] = upcomingArray[i][5];   //interviewee email
                upcomingArrayTight[j][5] = upcomingArray[i][13];  //spreadsheet row (for writing to sheet)
                doneRequestArray[j] = upcomingArray[i];
                j++;
            }
        }
        String[] titlesTight = new String[] {"Host Name","Host Emails","Hosting Date","Interviewee Name","Interviewee Email",""};
        JTable jtable = new JTable(upcomingArrayTight,titlesTight) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return column != 5;
            }
        };
        conformFiveColumnTable(jtable);
        jtable.setComponentPopupMenu(new TablePopupMenu(jtable));
        return jtable;
    }

    /**
     * <p>Constructs a JTable containing the info of interviewees' requests for hosts that have been moved to the
     * ignore list.</p>
     * <p>The JTable contains 4 columns for 1) hosting date, 2) interviewee name, 3) interviewee email, and
     * 4) Google Sheet row number. The last column is minimized and uneditable, and a custom
     * <code>TableRowSorter</code> is applied to the hosting date column - this is done by calling
     * <code>conformThreeColumnTable(...)</code>. The JTable is also constructed such that a customized JPopupMenu is
     * activated upon the JTable being right-clicked.</p>
     * @return the JTable for host requests that have been moved to the ignore list
     */
    public JTable getIgnoredJTable(){
        Object[][] ignoredArray = new Object[values.size()][values.get(0).size()];
        int ctr = 0;
        for(int i = 0; i < values.size(); i++){
            if(values.get(i).get(12).equals("3")){
                ctr++;
            }
            for(int j = 0; j < values.get(i).size(); j++){
                ignoredArray[i][j] = values.get(i).get(j);
            }
        }
        Object[][] ignoredArrayTight = new Object[ctr][4];
        ignoredRequestArray = new Object[ctr][13];
        int j = 0;
        for(int i = 0; i < ignoredArray.length; i++){
            if(ignoredArray[i][12].equals("3")) {
                ignoredArrayTight[j][0] = ignoredArray[i][4];   //hosting night
                ignoredArrayTight[j][1] = ignoredArray[i][1];   //interviewee name
                ignoredArrayTight[j][2] = ignoredArray[i][5];   //interviewee email
                ignoredArrayTight[j][3] = ignoredArray[i][13];  //spreadsheet row (for writing to sheet)
                ignoredRequestArray[j] = ignoredArray[i];
                j++;
            }
        }
        String[] titlesTight = new String[] {"Hosting Date","Interviewee Name","Interviewee Email","#"};
        JTable jtable = new JTable(ignoredArrayTight,titlesTight)  {
            @Override
            public boolean isCellEditable(int row, int column) {
                return column != 3;
            }
        };
        conformThreeColumnTable(jtable);
        jtable.setComponentPopupMenu(new TablePopupMenu(jtable));
        return jtable;
    }

    /**
     * <p>Sets the <code>DateComparator</code> for the hosting date column in the receipt/ignored tables, and also
     * minimizes and makes uneditable the final column containing the sheet row.</p>
     * @param jtable the receipt table
     */
    protected void conformThreeColumnTable(JTable jtable){
        jtable.setAutoCreateRowSorter(true);
        TableRowSorter trs = new TableRowSorter(jtable.getModel());
        trs.setComparator(0,new DateComparator());
        jtable.setRowSorter(trs);
        jtable.getColumnModel().getColumn(3).setMaxWidth(0);
    }

    /**
     * <p>Sets the <code>DateComparator</code> for the hosting date column in the pairing/done tables, and also
     * minimizes and makes uneditable the final column containing the sheet row.</p>
     * @param jtable the pairing table
     */
    protected void conformFiveColumnTable(JTable jtable){
        jtable.setAutoCreateRowSorter(true);
        TableRowSorter trs = new TableRowSorter(jtable.getModel());
        trs.setComparator(2,new DateComparator());
        jtable.setRowSorter(trs);
        jtable.getColumnModel().getColumn(5).setMaxWidth(0);
    }

    /**
     * <p>Returns the corresponding Google Sheet value given a row and column in the <code>TableModel</code> of the
     * receipt <code>JTable</code>. The "corresponding" value is the value of a table column (hosting date, interviewee
     * name, interviewee email) for a particular interviewee in the Google Sheet.</p>
     * <p>WARNING: To minimize lag, this method does not connect to the Google Drive and read values from the Google
     * Sheet. Instead, the value is returned from the instance variable <code>upcomingReceiptArray</code>. To ensure
     * the integrity of the values in <code>upcomingReceiptArray</code>, the method <code>updateValues()</code> should
     * be called to download values from the Google Sheet, followed by <code>updateUpcoming()</code> to refresh the
     * array of interviewees whose interview date hasn't passed yet (<code>upcomingArray</code>), followed by
     * <code>getReceiptJTable()</code> to refresh the array of specifically upcoming interviewees in wait of a receipt
     * email (<code>upcomingReceiptArray</code>), all before calling this method. Though this may seem cumbersome,
     * these methods are generally called in this order during the normal initialization and operation of the
     * <code>SheetsHandler</code>.</p>
     * @param tableModelRow the row number in the receipt table's <code>TableModel</code>
     * @param tableModelCol the column number in the receipt table's <code>TableModel</code>
     * @return the contents of the corresponding cell in the Google Sheet, which was stored in the array
     * <code>upcomingReceiptArray</code>
     */
    public Object getReceiptTableValueFromSheet(int tableModelRow, int tableModelCol){
        int sheetCol;
        switch(tableModelCol){
            case 0: sheetCol=4; break;
            case 1: sheetCol=1; break;
            case 2: sheetCol=5; break;
            default: throw new IllegalArgumentException("tableModelCol index "+tableModelCol);
        }
        return upcomingReceiptArray[tableModelRow][sheetCol];
    }

    /**
     * <p>Returns the corresponding Google Sheet value given a row and column in the <code>TableModel</code> of the
     * pairing <code>JTable</code>. The "corresponding" value is the value of a table column (host name, host email,
     * hosting date, interviewee name, interviewee email) for a particular interviewee in the Google Sheet.</p>
     * <p>WARNING: To minimize lag, this method does not connect to the Google Drive and read values from the Google
     * Sheet. Instead, the value is returned from the instance variable <code>upcomingPairingArray</code>. To ensure
     * the integrity of the values in <code>upcomingPairingArray</code>, the method <code>updateValues()</code> should
     * be called to download values from the Google Sheet, followed by <code>updateUpcoming()</code> to refresh the
     * array of interviewees whose interview date hasn't passed yet (<code>upcomingArray</code>), followed by
     * <code>getPairingJTable()</code> to refresh the array of specifically upcoming interviewees in wait of a pairing
     * email (<code>upcomingPairingArray</code>), all before calling this method. Though this may seem cumbersome,
     * these methods are generally called in this order during the normal initialization and operation of the
     * <code>SheetsHandler</code>.</p>
     * @param tableModelRow the row number in the pairing table's <code>TableModel</code>
     * @param tableModelCol the column number in the pairing table's <code>TableModel</code>
     * @return the contents of the corresponding cell in the Google Sheet, which was stored in the array
     * <code>upcomingPairingArray</code>
     */
    public Object getPairingTableValueFromSheet(int tableModelRow, int tableModelCol){
        int sheetCol;
        switch(tableModelCol){
            case 0: sheetCol=0; break;
            case 1: sheetCol=11; break;
            case 2: sheetCol=4; break;
            case 3: sheetCol=1; break;
            case 4: sheetCol=5; break;
            default: throw new IllegalArgumentException("tableModelCol index "+tableModelCol);
        }
        return upcomingPairingArray[tableModelRow][sheetCol];
    }

    /**
     * <p>Returns the corresponding Google Sheet value given a row and column in the <code>TableModel</code> of the
     * ignored <code>JTable</code>. The "corresponding" value is the value of a table column (hosting date, interviewee
     * name, interviewee email) for a particular interviewee in the Google Sheet.</p>
     * <p>WARNING: To minimize lag, this method does not connect to the Google Drive and read values from the Google
     * Sheet. Instead, the value is returned from the instance variable <code>ignoredRequestArray</code>. To ensure
     * the integrity of the values in <code>ignoredRequestArray</code>, the method <code>updateValues()</code> should
     * be called to download values from the Google Sheet, followed by <code>updateUpcoming()</code> to refresh the
     * array of interviewees whose interview date hasn't passed yet (<code>upcomingArray</code>), followed by
     * <code>getReceiptJTable()</code> to refresh the array of specifically upcoming interviewees whose hosting
     * requests are ignored (<code>ignoredRequestArray</code>), all before calling this method. Though this may
     * seem cumbersome, these methods are generally called in this order during the normal initialization and operation
     * of the <code>SheetsHandler</code>.</p>
     * @param tableModelRow the row number in the receipt table's <code>TableModel</code>
     * @param tableModelCol the column number in the receipt table's <code>TableModel</code>
     * @return the contents of the corresponding cell in the Google Sheet, which was stored in the array
     * <code>ignoredRequestArray</code>
     */
    public Object getIgnoredTableValueFromSheet(int tableModelRow, int tableModelCol){
        int sheetCol;
        switch(tableModelCol){
            case 0: sheetCol=4; break;
            case 1: sheetCol=1; break;
            case 2: sheetCol=5; break;
            default: throw new IllegalArgumentException("tableModelCol index "+tableModelCol);
        }
        return ignoredRequestArray[tableModelRow][sheetCol];
    }

    /**
     * <p>Returns the corresponding Google Sheet value given a row and column in the <code>TableModel</code> of the
     * "done" <code>JTable</code>. The "corresponding" value is the value of a table column (host name, host email,
     * hosting date, interviewee name, interviewee email) for a particular interviewee in the Google Sheet.</p>
     * <p>WARNING: To minimize lag, this method does not connect to the Google Drive and read values from the Google
     * Sheet. Instead, the value is returned from the instance variable <code>upcomingPairingArray</code>. To ensure
     * the integrity of the values in <code>upcomingPairingArray</code>, the method <code>updateValues()</code> should
     * be called to download values from the Google Sheet, followed by <code>updateUpcoming()</code> to refresh the
     * array of interviewees whose interview date hasn't passed yet (<code>upcomingArray</code>), followed by
     * <code>getDoneJTable()</code> to refresh the array of specifically upcoming interviewees in wait of a pairing
     * email (<code>upcomingPairingArray</code>), all before calling this method. Though this may seem cumbersome,
     * these methods are generally called in this order during the normal initialization and operation of the
     * <code>SheetsHandler</code>.</p>
     * @param tableModelRow the row number in the pairing table's <code>TableModel</code>
     * @param tableModelCol the column number in the pairing table's <code>TableModel</code>
     * @return the contents of the corresponding cell in the Google Sheet, which was stored in the array
     * <code>upcomingPairingArray</code>
     */
    public Object getDoneTableValueFromSheet(int tableModelRow, int tableModelCol){
        int sheetCol;
        switch(tableModelCol){
            case 0: sheetCol=0; break;
            case 1: sheetCol=11; break;
            case 2: sheetCol=4; break;
            case 3: sheetCol=1; break;
            case 4: sheetCol=5; break;
            default: throw new IllegalArgumentException("tableModelCol index "+tableModelCol);
        }
        return doneRequestArray[tableModelRow][sheetCol];
    }

    /**
     * <p>Returns the ID of the cell in the Google Sheet corresponding to the value at the specified row and column in
     * the <code>TableModel</code> of the receipt <code>JTable</code>. This is useful for getting the cell ID when
     * updating a value in the Google Sheet.</p>
     * @param tableModelRow the row number of the specified cell in the receipt table's <code>TableModel</code>
     * @param tableModelCol the column number of the specified cell in the receipt table's <code>TableModel</code>
     * @return the cell ID of the corresponding value in the Google Sheet (i.e. "Sheet2!B7")
     */
    public String getReceiptTableValueSheetRange(int tableModelRow, int tableModelCol){
        String range = "Sheet2!";
        int sheetCol;
        switch(tableModelCol){
            case 0: sheetCol=4; break;
            case 1: sheetCol=1; break;
            case 2: sheetCol=5; break;
            default: throw new IllegalArgumentException("tableModelCol index "+tableModelCol);
        }
        range += indexToLetter[sheetCol];
        range += upcomingReceiptArray[tableModelRow][13];
        return range;
    }

    /**
     * <p>Returns the ID of the cell in the Google Sheet corresponding to the value at the specified row and column in
     * the <code>TableModel</code> of the pairing <code>JTable</code>. This is useful for getting the cell ID when
     * updating a value in the Google Sheet.</p>
     * @param tableModelRow the row number of the specified cell in the receipt table's <code>TableModel</code>
     * @param tableModelCol the column number of the specified cell in the receipt table's <code>TableModel</code>
     * @return the cell ID of the corresponding value in the Google Sheet (i.e. "Sheet2!B7")
     */
    public String getPairingTableValueSheetRange(int tableModelRow, int tableModelCol){
        String range = "Sheet2!";
        int sheetCol;
        switch(tableModelCol){
            case 0: sheetCol=0; break;
            case 1: sheetCol=11; break;
            case 2: sheetCol=4; break;
            case 3: sheetCol=1; break;
            case 4: sheetCol=5; break;
            default: throw new IllegalArgumentException("tableModelCol index "+tableModelCol);
        }
        range += indexToLetter[sheetCol];
        range += upcomingPairingArray[tableModelRow][13];
        return range;
    }

    /**
     * <p>Returns the ID of the cell in the Google Sheet corresponding to the value at the specified row and column in
     * the <code>TableModel</code> of the ignored <code>JTable</code>. This is useful for getting the cell ID when
     * updating a value in the Google Sheet.</p>
     * @param tableModelRow the row number of the specified cell in the ignored table's <code>TableModel</code>
     * @param tableModelCol the column number of the specified cell in the ignored table's <code>TableModel</code>
     * @return the cell ID of the corresponding value in the Google Sheet (i.e. "Sheet2!B7")
     */
    public String getIgnoredTableValueSheetRange(int tableModelRow, int tableModelCol){
        String range = "Sheet2!";
        int sheetCol;
        switch(tableModelCol){
            case 0: sheetCol=4; break;
            case 1: sheetCol=1; break;
            case 2: sheetCol=5; break;
            default: throw new IllegalArgumentException("tableModelCol index "+tableModelCol);
        }
        range += indexToLetter[sheetCol];
        range += ignoredRequestArray[tableModelRow][13];
        return range;
    }

    /**
     * <p>Returns the ID of the cell in the Google Sheet corresponding to the value at the specified row and column in
     * the <code>TableModel</code> of the done <code>JTable</code>. This is useful for getting the cell ID when
     * updating a value in the Google Sheet.</p>
     * @param tableModelRow the row number of the specified cell in the done table's <code>TableModel</code>
     * @param tableModelCol the column number of the specified cell in the done table's <code>TableModel</code>
     * @return the cell ID of the corresponding value in the Google Sheet (i.e. "Sheet2!B7")
     */
    public String getDoneTableValueSheetRange(int tableModelRow, int tableModelCol){
        String range = "Sheet2!";
        int sheetCol;
        switch(tableModelCol){
            case 0: sheetCol=0; break;
            case 1: sheetCol=11; break;
            case 2: sheetCol=4; break;
            case 3: sheetCol=1; break;
            case 4: sheetCol=5; break;
            default: throw new IllegalArgumentException("tableModelCol index "+tableModelCol);
        }
        range += indexToLetter[sheetCol];
        range += doneRequestArray[tableModelRow][13];
        return range;
    }
    /**
     * <p>Accessor method for the plea email signature (i.e. the names of all the hosting coordinators). This
     * replaces the Merge Tag "[PLEA SIGNATURE]" in plea emails.</p>
     * @return the plea email signature (i.e. the names of all the hosting coordinators)
     */
    public String getPleaEmailSignature(){
        return pleaEmailSignature;
    }

    /**
     * <p>Accessor method for the receipt email body, which contains Merge Tags. The Merge Tags are</p>
     * <ol>
     *     <li>[INTERVIEWEE NAME]</li>
     *     <li>[HOSTING DATE]</li>
     *     <li>[PREFERENCE TABLE]</li>
     *     <li>[SIGNATURE]</li>
     * </ol>
     * @return the draft receipt email body
     */
    public String getReceiptEmailText(){
        return receiptEmailText;
    }

    /**
     * <p>Accessor method for the pairing email body, which contains Merge Tags. The Merge Tags are</p>
     * <ol>
     *     <li>[INTERVIEWEE NAME]</li>
     *     <li>[HOSTING DATE]</li>
     *     <li>[HOST NAME]</li>
     *     <li>[SIGNATURE]</li>
     * </ol>
     * @return the draft pairing email body
     */
    public String getPairingEmailText(){
        return pairingEmailText;
    }

    /**
     * <p>Accessor method for the plea email body, which contains Merge Tags. The Merge Tags are</p>
     * <ol>
     *     <li>[INTERVIEWEE TABLE]</li>
     *     <li>[PLEA SIGNATURE]</li>
     * </ol>
     * @return the draft plea email body
     */
    public String getPleaEmailText(){
        return pleaEmailText;
    }

    /**
     * <p>Sets the contents of the interviewee info area (at the right side of the window) to show detailed
     * info about the interviewee selected in the table. Unlike info in the table, the info displayed in the
     * text area is not editable through the application.</p>
     * @param jTable the table whose currently selected interviewee will have their details shown in the text area
     * @param jTextArea the text area whose contents will be set to show detailed info about the interviewee
     * @param tableName the name of the table, which can be any one of the following constants:
     *                  <ul>
     *                  <li><code>PritzkerEHR.CONFIRM_INT_REQUEST_RECEIVED</code></li>
     *                  <li><code>PritzkerEHR.CONFIRM_HOST_TO_INTERVIEWEES</code></li>
     *                  <li><code>PritzkerEHR.VIEW_IGNORE_LIST</code></li>
     *                  </ul>
     */
    public void updateInfoArea(JTable jTable, JTextArea jTextArea, String tableName){
        int i = jTable.getSelectedRow();
        if(i == -1){  //no selection
            return;
        }
        int ind = jTable.convertRowIndexToModel(i);
        String info = "";
        if(tableName.equals(PritzkerEHR.CONFIRM_INT_REQUEST_RECEIVED)) {  //confirmation table
            info = "Name: " + upcomingReceiptArray[ind][1] + "\n" +
                    "Gender: " + upcomingReceiptArray[ind][2] + "\n" +
                    "Interview Date: " + upcomingReceiptArray[ind][3] + "\n" +
                    "Hosting Date: " + upcomingReceiptArray[ind][4] + "\n" +
                    "Undergraduate: " + upcomingReceiptArray[ind][7] + "\n" +
                    "Preferred Host Gender: " + upcomingReceiptArray[ind][8] + "\n" +
                    "Allergies: " + upcomingReceiptArray[ind][9] + "\n" +
                    "Interest Groups: " + upcomingReceiptArray[ind][10] + "\n" +
                    "\nCONTACT INFO:\n" +
                    "Email Address: " + upcomingReceiptArray[ind][5] + "\n" +
                    "Phone Number: " + upcomingReceiptArray[ind][6];
        }
        else if(tableName.equals(PritzkerEHR.CONFIRM_HOST_TO_INTERVIEWEES)){
            info = "Name: " + upcomingPairingArray[ind][1] + "\n" +
                    "Gender: " + upcomingPairingArray[ind][2] + "\n" +
                    "Interview Date: " + upcomingPairingArray[ind][3] + "\n" +
                    "Hosting Date: " + upcomingPairingArray[ind][4] + "\n" +
                    "Undergraduate: " + upcomingPairingArray[ind][7] + "\n" +
                    "Preferred Host Gender: " + upcomingPairingArray[ind][8] + "\n" +
                    "Allergies: " + upcomingPairingArray[ind][9] + "\n" +
                    "Interest Groups: " + upcomingPairingArray[ind][10] + "\n" +
                    "\nCONTACT INFO:\n" +
                    "Email Address: " + upcomingPairingArray[ind][5] + "\n" +
                    "Phone Number: " + upcomingPairingArray[ind][6];
        }
        else if(tableName.equals(PritzkerEHR.VIEW_IGNORE_LIST)){
            info = "Name: " + ignoredRequestArray[ind][1] + "\n" +
                    "Gender: " + ignoredRequestArray[ind][2] + "\n" +
                    "Interview Date: " + ignoredRequestArray[ind][3] + "\n" +
                    "Hosting Date: " + ignoredRequestArray[ind][4] + "\n" +
                    "Undergraduate: " + ignoredRequestArray[ind][7] + "\n" +
                    "Preferred Host Gender: " + ignoredRequestArray[ind][8] + "\n" +
                    "Allergies: " + ignoredRequestArray[ind][9] + "\n" +
                    "Interest Groups: " + ignoredRequestArray[ind][10] + "\n" +
                    "\nCONTACT INFO:\n" +
                    "Email Address: " + ignoredRequestArray[ind][5] + "\n" +
                    "Phone Number: " + ignoredRequestArray[ind][6];
        }
        else if(tableName.equals(PritzkerEHR.VIEW_COMPLETED_REQUESTS)) {
            info = "Name: " + doneRequestArray[ind][1] + "\n" +
                    "Gender: " + doneRequestArray[ind][2] + "\n" +
                    "Interview Date: " + doneRequestArray[ind][3] + "\n" +
                    "Hosting Date: " + doneRequestArray[ind][4] + "\n" +
                    "Undergraduate: " + doneRequestArray[ind][7] + "\n" +
                    "Preferred Host Gender: " + doneRequestArray[ind][8] + "\n" +
                    "Allergies: " + doneRequestArray[ind][9] + "\n" +
                    "Interest Groups: " + doneRequestArray[ind][10] + "\n" +
                    "\nCONTACT INFO:\n" +
                    "Email Address: " + doneRequestArray[ind][5] + "\n" +
                    "Phone Number: " + doneRequestArray[ind][6];
        }
        jTextArea.setText(info);
        jTextArea.setCaretPosition(0);
    }

    /**
     * <p>Accessor method for getting the instance variable <code>upcomingReceiptArray</code></p>
     * @return upcomingReceiptArray
     */
    protected Object[][] getUpcomingReceiptArray(){
        return upcomingReceiptArray;
    }

    /**
     * <p>Accessor method for getting the instance variable <code>upcomingPairingArray</code></p>
     * @return upcomingPairingArray
     */
    protected Object[][] getUpcomingPairingArray(){
        return upcomingPairingArray;
    }

    /**
     * <p>Accessor method for getting the instance variable <code>ignoredRequestArray</code></p>
     * @return ignoredRequestArray
     */
    protected Object[][] getIgnoredRequestArray(){
        return ignoredRequestArray;
    }

    /**
     * <p>Accessor method for getting the instance variable <code>ignoredRequestArray</code></p>
     * @return ignoredRequestArray
     */
    protected Object[][] getDoneRequestArray(){
        return doneRequestArray;
    }

    /**
     * <p>Generates a <code>String</code> representing a 2 column HTML table containing the titles of the hosting
     * preferences in the lefthand column and the interviewee's hosting preferences in the righthand column. This
     * table can then be included in a receipt email, in which it will replace the merge tag '[PREFERENCE TABLE]'."</p>
     * <p>The intervieweeData array should be one of the four interviewee arrays obtained through the
     * <code>protected</code> methods with the title 'get[X]Array()'.</p>
     * @param intervieweeData the interviewee data array. Should be obtained before being passed as an argument by
     *                       by calling one of the following methods:
     *                        <ul>
     *                          <li>getUpcomingReceiptArray()</li>
     *                          <li>getUpcomingPairingArray()</li>
     *                          <li>getIgnoredRequestArray()</li>
     *                          <li>getDoneRequestArray()</li>
     *                        </ul>
     * @param row the row number in <code>intervieweeData</code> corresponding to the relevant interviewee
     * @return
     */
    protected String getHTMLPreferenceTable(Object[][] intervieweeData, int row){
        StringBuilder sb = new StringBuilder("<table>");
        for(int i = 0; i < fieldNames.length; i++){
            sb.append("<tr><td><b>"+fieldNames[i]+"</b></td><td>"+intervieweeData[row][i+1]+"</td></tr>");
        }
        sb.append("</table>");
        return sb.toString();
    }

    /**
     * <p>Writes the specified <code>input</code> to a single cell specified by <code>range</code>.
     * @param range the cell ID in the Sheet2 of the Google Sheet (i.e. "B2")
     * @param input the contents to be written to the cell (ideally a <code>String</code>)
     */
    public void write(String range, Object input){
        Sheets service = null;
        ValueRange vr = new ValueRange();
        vr.setRange(range);
        vr.setValues(Arrays.asList(Arrays.asList(input)));
        Sheets.Spreadsheets.Values.Update request = null;
        try {
            service = getSheetsService();
            request = service.spreadsheets().values().update(spreadsheetID, range, vr);
            request.setValueInputOption("USER_ENTERED");
            UpdateValuesResponse response = request.execute();
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(pritzkerEHR,"Pritzker EHR has just had trouble connecting to Google Drive. Please wait two seconds and try again.","Oops!",JOptionPane.WARNING_MESSAGE,pritzkerEHR.getWindowIcon());
        }
    }

    /**
     * <p>A convenience method with public access for executing several methods in order to properly download
     * interviewee info and update the GUI components of the associated instance of PritzkerEHR. The methods
     * executed are as follows in the following order:</p>
     * <ol>
     *     <li><code>updateValues()</code></li>
     *     <li><code>updateSheet2()</code></li>
     *     <li><code>updateUpcoming()</code></li>
     *     <li><code>updatePleaAndTables()</code></li>
     * </ol>
     */
    public void refresh(){
        updateValues();
        updateSheet2();
        updateUpcoming();
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                pritzkerEHR.updatePleaAndTables();
            }
        });
    }

    /**     PRIVATE     **/
    /**
     * <p>Creates an authorized Credential object.</p>
     * @return an authorized Credential object.
     * @throws IOException
     */
    private Credential authorize() throws IOException {
        // Load client secrets.
        InputStream in = SheetsHandler.class.getResourceAsStream("client_secret.json");
        GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));
        // Build flow and trigger user authorization request.
        GoogleAuthorizationCodeFlow flow =
                new GoogleAuthorizationCodeFlow.Builder(
                        HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
                        .setDataStoreFactory(DATA_STORE_FACTORY)
                        .setAccessType("offline")
                        .build();
        Credential credential = new AuthorizationCodeInstalledApp(
                flow, new LocalServerReceiver()).authorize("user");
        System.out.println(
                "Credentials saved to " + DATA_STORE_DIR.getAbsolutePath());
        return credential;
    }

    /**
     * <p>Build and return an authorized Sheets API client service.</p>
     * @return an authorized Sheets API client service
     * @throws IOException
     */
    private Sheets getSheetsService() throws IOException {
        return new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, credential)
                .setApplicationName(APPLICATION_NAME)
                .build();
    }

    /**
     * <p>Reads the contents of Sheet1 into the <code>SheetsHandler</code>.
     * The contents are stored in the instance variable <code>values</code>.</p>
     */
    private void updateValues(){
        Sheets service = null;
        values = null;
        try {
            service = getSheetsService();
            String range = "Sheet2!A:M";
            ValueRange response = service.spreadsheets().values().get(spreadsheetID, range).execute();
            values = response.getValues();
        } catch (Throwable t) {
            t.printStackTrace();
            JOptionPane.showMessageDialog(pritzkerEHR,"Pritzker EHR has just had trouble connecting to Google Drive. Please wait two seconds and try again.","Oops!",JOptionPane.WARNING_MESSAGE,pritzkerEHR.getWindowIcon());
            System.exit(1);
        }
        //add row #
        for(int i = 0; i < values.size(); i++){
            values.get(i).add(i+1);
            //???? why i+1 and not i+2?
            System.out.println(values.get(i));
        }
    }

    /**
     * <p>Updates list of upcoming interviewees based on data in Sheet2.
     * To ensure that Sheet2 is updated, updateValues() should always be called first.</p>
     */
    private void updateUpcoming(){
        Object[][] allData = new Object[values.size()][values.get(0).size()];
        for(int i = 0; i < values.size(); i++){
            for(int j = 0; j < values.get(i).size(); j++){
                allData[i][j] = values.get(i).get(j);
            }
        }
        upcomingGuests = new ArrayList<ArrayList<Object>>();
        Date today = new Date();
        Date arg = null;
        SimpleDateFormat format = new SimpleDateFormat("MM/dd/yyyy");
        boolean suppressWarnings = false;
        int warningCount = 0;
        for(int i = 1; i < allData.length; i++){
            try{
                arg = format.parse((String)(allData[i][3]));
            }
            catch (ParseException e){
                if(!suppressWarnings) {
                    JOptionPane.showMessageDialog(pritzkerEHR, (String) (allData[i][3]) + " is not a parseable date in MM/DD/YYYY format", "Pritzker EHR says:", JOptionPane.WARNING_MESSAGE, pritzkerEHR.getWindowIcon());
                    e.printStackTrace();
                    warningCount++;
                    if (warningCount % 3 == 0) {
                        int choice = JOptionPane.showConfirmDialog(pritzkerEHR, "Suppress future unparseable date warnings?", "Pritzker EHR says:", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE, pritzkerEHR.getWindowIcon());
                        if (choice == JOptionPane.YES_OPTION) {
                            suppressWarnings = true;
                        }
                    }
                }

            }
            //allData[i][3]
//            if(arg.compareTo(today) >= 0){
//                upcomingGuests.add((ArrayList)values.get(i));
//            }
            //upcoming guests == values. It was decided in version 1.3 that this caused problems when people misentered their hosting dates
            upcomingGuests.add((ArrayList)values.get(i));
        }
    }

    /**
     * <p>Copies new entries from Sheet1 into Sheet2. Once copied, entries on Sheet1 are marked with a "1" in column L.
     */
    private void updateSheet2(){
        ArrayList<Integer> newArrivals = new ArrayList<Integer>();
        Sheets service = null;
        List<List<Object>> rawValues = null;
        try {
            service = getSheetsService();
            String range = "A:L";
            ValueRange response = service.spreadsheets().values().get(spreadsheetID, range).execute();
            rawValues = response.getValues();
        } catch (Throwable t) {
            t.printStackTrace();
            JOptionPane.showMessageDialog(pritzkerEHR,"Pritzker EHR has just had trouble connecting to Google Drive. Please wait two seconds and try again.","Oops!",JOptionPane.WARNING_MESSAGE,pritzkerEHR.getWindowIcon());
            System.exit(1);
        }
        for(int i = 0; i < rawValues.size(); i++){
            if(rawValues.get(i).size() < 12){
                //new arrival
                rawValues.get(i).set(0,"");
                while(rawValues.get(i).size() < 12){
                    rawValues.get(i).add("");
                }
                rawValues.get(i).add("0");
                newArrivals.add(i);
            }
        }
        for(int i = 0; i < newArrivals.size(); i++){
            String range = "Sheet2!A"+(values.size()+i+1)+":M"+(values.size()+i+1);
            ValueRange vr = new ValueRange();
            vr.setRange(range);
            ArrayList<List<Object>> row = (new ArrayList<List<Object>>());
            row.add(rawValues.get(newArrivals.get(i)));
            vr.setValues(row);
            Sheets.Spreadsheets.Values.Update request = null;
            try {
                request = service.spreadsheets().values().update(spreadsheetID, range, vr);
                request.setValueInputOption("USER_ENTERED");
                UpdateValuesResponse response = request.execute();
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(pritzkerEHR,"Pritzker EHR has just had trouble connecting to Google Drive. Please wait two seconds and try again.","Oops!",JOptionPane.WARNING_MESSAGE,pritzkerEHR.getWindowIcon());
                System.exit(1);
            }
            range = "L"+(newArrivals.get(i)+1);
            List<List<Object>> val = new ArrayList<List<Object>>();
            val.add(new ArrayList<Object>());
            val.get(0).add("1");
            vr = new ValueRange();
            vr.setRange(range);
            vr.setValues(val);
            try {
                request = service.spreadsheets().values().update(spreadsheetID, range, vr);
                request.setValueInputOption("USER_ENTERED");
                UpdateValuesResponse response = request.execute();
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(pritzkerEHR,"Pritzker EHR has just had trouble connecting to Google Drive. Please wait two seconds and try again.","Oops!",JOptionPane.WARNING_MESSAGE,pritzkerEHR.getWindowIcon());
                System.exit(1);
            }
        }
        updateValues();
        updateUpcoming();
    }

    /**
     * <p>Sorts the interviewee info array by hosting date (should be column 4 but can be specified otherwise).</p>
     * @param array the interviewee info array to be sorted
     * @param columnNbr the number of the column containing dates (should be 4 but can be specified otherwise)
     */
    private void sortArrayByDate(Object[][] array, int columnNbr){
        Arrays.sort(array, new MultidimensionalDateComparator(columnNbr));
    }

    /**
     * <p>Reads program presets from the Google Sheet "PritzkerHosting_PATH." This method should always be called
     * at construction to ensure that the program is reading from the intended Google Sheet containing interviewee info.
     * </p>
     * <p>The presets are as follows:</p>
     * <ol>
     *     <li>the spreadsheet ID of the Google Sheet currently being used to store interviewee host requests</li>
     *     <li>the email signature used to sign plea emails (i.e. all the names of the hosting coordinators)</li>
     *     <li>the draft email body for receipt emails, containing Merge Tags</li>
     *     <li>the draft email body for pairing emails, containing Merge Tags</li>
     *     <li>the draft email body for plea emails, containing Merge Tags</li>
     * </ol>
     */
    private void loadPATH(){
        Sheets service = null;
        List<List<Object>> cellValues = null;
        try {
            service = getSheetsService();
            String range = "B1:B5";
            ValueRange response = service.spreadsheets().values().get(PATH_ID, range).execute();
            cellValues = response.getValues();
        } catch (Throwable t) {
            t.printStackTrace();
            JOptionPane.showMessageDialog(pritzkerEHR,"Pritzker EHR has just had trouble connecting to Google Drive. Please wait two seconds and try again.","Oops!",JOptionPane.WARNING_MESSAGE,pritzkerEHR.getWindowIcon());
            System.exit(1);
        }
        spreadsheetID = cellValues.get(0).get(0).toString();
        pleaEmailSignature = cellValues.get(1).get(0).toString();
        receiptEmailText = cellValues.get(2).get(0).toString();
        pairingEmailText = cellValues.get(3).get(0).toString();
        pleaEmailText = cellValues.get(4).get(0).toString();
    }

    /**
     * <p>Opens a <code>JOptionPane</code> that allows the user to manually edit an interviewee's preferences. The
     * <code>JOptionPane</code> contains several <code>JTextField</code>s that are auto-populated with the interviewee's
     * preferences. The user can then edit the contents of the <code>JTextField</code>s and click the "OK" button at
     * the bottom of the <code>JOptionPane</code> to save the edits. Once saved, the edits cannot be undone, though the
     * interviewee's original preferences will always be kept unaltered in the first sheet of the Google Doc.</p>
     * @param sheetRow the row number of the interviewee in Sheet2 of the Google Doc
     */
    private void openIntervieweeEditor(int sheetRow) {
        JTextField[] fields = new JTextField[10];
        JPanel main = new JPanel(new GridLayout(10,2));
        for(int i=0; i<fields.length; i++){
            fields[i] = new JTextField(values.get(sheetRow-1).get(i+1).toString(),20);
            JLabel jLabel = new JLabel(fieldNames[i]+":");
            main.add(jLabel);
            main.add(fields[i]);
        }
        int choice = JOptionPane.showConfirmDialog(pritzkerEHR,main,"Edit Interviewee Info",JOptionPane.OK_CANCEL_OPTION,JOptionPane.PLAIN_MESSAGE);
        if(choice == JOptionPane.OK_OPTION){
            for(int i = 0; i < fields.length; i++) {
                if(!fields[i].getText().equals(values.get(sheetRow-1).get(i+1).toString())) {
                    write("Sheet2!"+indexToLetter[i+1]+sheetRow,fields[i].getText());
                }
            }
            JOptionPane.showMessageDialog(pritzkerEHR,"All requested edits finished!","Pritzker EHR says:",JOptionPane.INFORMATION_MESSAGE);
        }
    }

    /**
     * <p>The popup menu class for the interviewee <code>JTable</code>s.Each instance of <code>TablePopupMenu</code>
     * is associated with a single interviewee <code>JTable</code>.</p>
     * <p><code>TablePopupMenu</code> has 4 options that allow the user to manually edit the hosting status
     * of a selected interviewee to any one of 1) awaiting receipt of request, 2) awaiting pairing email,
     * 3) already paired with host, or 4) host request ignored. The interviewee <code>JTable</code>s are configured such
     * that the <code>TablePopupMenu</code> is made visible upon the <code>JTable</code> being right-clicked.
     * The <code>TablePopupMenu</code>, once made visible, is shown at the location of the right-click, which is
     * necessarily near the selected interviewee's name in the table.</p>
     */
    private class TablePopupMenu extends JPopupMenu implements ActionListener, PopupMenuListener {
        private JTable jTable;
        private JMenuItem refreshAll;
        private JMenuItem moveTo0;
        private JMenuItem moveTo1;
        private JMenuItem moveTo2;
        private JMenuItem moveTo3;
        private JMenuItem edit;

        /**
         * <p>Constructor for the <code>TablePopupMenu</code>.</p>
         * <p>The <code>JTable</code> passed as an argument is stored as an instance variable
         * so that the <code>TablePopupMenu</code> can find the currently selected table tow
         * and thus find the currently selected interviewee whose hosting status is to be
         * modified by the user.</p>
         *
         * @param jTable the interviewee table associated with the <code>TablePopupMenu</code>
         */
        public TablePopupMenu(JTable jTable){
            this.jTable = jTable;
            addPopupMenuListener(this);
            refreshAll = new JMenuItem("Refresh Tables with Data from Docs");
            refreshAll.setIcon(new ImageIcon(getClass().getResource("resources/refresh_small.gif")));
            refreshAll.addActionListener(this);
            ImageIcon icon0 = null, icon1 = null, icon2 = null, icon3 = null, icon4 = null;
            try {
                icon0 = new ImageIcon(ImageIO.read(getClass().getResource("resources/shrug_icon_small.png")));
                icon1 = new ImageIcon(ImageIO.read(getClass().getResource("resources/pray_icon_small.png")));
                icon2 = new ImageIcon(ImageIO.read(getClass().getResource("resources/family_icon_small.png")));
                icon3 = new ImageIcon(ImageIO.read(getClass().getResource("resources/ghost_icon_small.png")));
                icon4 = new ImageIcon(ImageIO.read(getClass().getResource("resources/edit_small.gif")));
            } catch (IOException e) {
                e.printStackTrace();
            }
            moveTo0 = new JMenuItem("Move to Awaiting Confirmation", icon0);
            moveTo0.addActionListener(this);
            moveTo1 = new JMenuItem("Move to Awaiting Host Pairing", icon1);
            moveTo1.addActionListener(this);
            moveTo2 = new JMenuItem("Move to Already Paired with Host", icon2);
            moveTo2.addActionListener(this);
            moveTo3 = new JMenuItem("Move to Ignore/Duplicate list", icon3);
            moveTo3.addActionListener(this);
            edit = new JMenuItem("Edit Interviewee Preferences", icon4);
            edit.addActionListener(this);
            setPopupMenuContents();
        }

        /**
         *  <p>Sets the menu contents (either all menu options or just the refresh option) depending on whether the
         *  table contains any interviewees. This method is called each time the <code>TablePopupMenu</code> is
         *  opened, as it is called in <code>popupMenuWillBecomeVisible(final PopupMenuEvent e)</code></p>
         */
        public void setPopupMenuContents(){
            removeAll();
            add(refreshAll);
            if(jTable.getRowCount() > 0){
                addSeparator();
                add(moveTo0);
                add(moveTo1);
                add(moveTo2);
                add(moveTo3);
                add(edit);
            }
        }

        /**
         * <p>The handler method for <code>ActionEvent</code>s fired by clicking
         * one of the popup menu items.</p>
         * @param ae the <code>ActionEvent</code> being handled
         */
        public void actionPerformed(ActionEvent ae){
            Object source = ae.getSource();
            if(jTable.getRowCount() > 0) {
                int sheetRow = Integer.parseInt(jTable.getValueAt(jTable.getSelectedRow(), jTable.convertColumnIndexToView(jTable.getColumnCount() - 1)).toString());
                if (source.equals(moveTo0)) {
                    write("Sheet2!M" + sheetRow, "0");
                } else if (source.equals(moveTo1)) {
                    write("Sheet2!M" + sheetRow, "1");
                } else if (source.equals(moveTo2)) {
                    write("Sheet2!M" + sheetRow, "2");
                } else if (source.equals(moveTo3)) {
                    write("Sheet2!M" + sheetRow, "3");
                }
                else if (source.equals(edit)) {
                    openIntervieweeEditor(sheetRow);
                }
            }
            refresh();
        }

        /**
         * <p>Selects the <code>JTable</code> row where the right-click happened.</p>
         * @param e
         */
        @Override
        public void popupMenuWillBecomeVisible(final PopupMenuEvent e) {
            SwingUtilities.invokeLater(new Runnable() {
                @Override
                public void run() {
                    setPopupMenuContents();
                    int rowAtPoint = jTable.rowAtPoint(SwingUtilities.convertPoint((JPopupMenu)(e.getSource()), new Point(0, 0), jTable));
                    if (rowAtPoint > -1) {
                        jTable.setRowSelectionInterval(rowAtPoint, rowAtPoint);
                    }
                }
            });
        }

        /**
         * <p>Stub method with no implementation.</p>
         * @param e
         */
        @Override
        public void popupMenuWillBecomeInvisible(PopupMenuEvent e) {
            // TODO Auto-generated method stub
        }

        /**
         * <p>Stub method with no implementation.</p>
         * @param e
         */
        @Override
        public void popupMenuCanceled(PopupMenuEvent e) {
            // TODO Auto-generated method stub
        }
    }

    /**
     * <p><code>Comparator</code> class for ordering calendar dates passed as <code>String</code>s in
     * MM/dd/yyyy format.</p>
     */
    private class DateComparator implements Comparator {
        /**
         * <p>The method for comparing two dates passed as <code>String</code>s in MM/dd/yyyy format.</p>
         * @param o1 date 1 (should be a <code>String</code> in MM/dd/yyyy format)
         * @param o2 date 2 (should be a <code>String</code> in MM/dd/yyyy format)
         * @return 1 if o1 is later than o2; -1 if o1 is earlier than o2; 0 if o1 equals o2
         */
        public int compare(Object o1, Object o2) {
            SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
            try {
                Date d1 = sdf.parse((String) o1);
                Date d2 = sdf.parse((String) o2);
                return d1.compareTo(d2);
            }
            catch(ParseException e){
                e.printStackTrace();
            }
            return 0;
        }
    }

    /**
     * <p><code>Comparator</code> class for ordering rows in a 2D Array by calendar date values in one of the columns.
     * </p>
     * <p>The column number containing dates used for ordering must be specified at construction.</p>
     */
    private class MultidimensionalDateComparator implements Comparator {
        /**
         * <p>The index of the column containing the dates used for sorting the rows.</p>
         */
        private int sortingColumnNumber = -1;

        /**
         * <p>Constructor for the <code>MultidimensionalDateComparator</code></p>
         * @param scn the index of the column containing dates used for sorting the rows
         */
        public MultidimensionalDateComparator(int scn){
            sortingColumnNumber = scn;
        }

        /**
         * <p>The method for comparing two 1D arrays based on an array index containing dates in MM/dd/yyyy format in both arrays.</p>
         * @param o1 ID array 1 (the element at index 'sortingColumnNumber' should be a <code>String</code> in MM/dd/yyyy format)
         * @param o2 1D array 2 (the element at index 'sortingColumnNumber' should be a <code>String</code> in MM/dd/yyyy format)
         * @return 1 if o1 is later than o2; -1 if o1 is earlier than o2; 0 if o1 equals o2
         */
        public int compare(Object o1, Object o2) {
            SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");
            try {
                Date d1 = sdf.parse(((Object[])o1)[sortingColumnNumber].toString());
                Date d2 = sdf.parse(((Object[])o2)[sortingColumnNumber].toString());
                return d1.compareTo(d2);
            }
            catch(ParseException e){
                e.printStackTrace();
            }
            return 0;
        }
    }
}
