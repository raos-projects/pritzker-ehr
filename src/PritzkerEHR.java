import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Properties;
import javax.imageio.ImageIO;
import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.swing.*;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;

/**
 * <p>Pritzker Electronic Hosting Record - a GUI for managing interviewee hosting requests.</p>
 *
 * <p>Created by Saieesh Rao, MD '21 beginning on 9/8/2017.
 * Code for boolean sendMail(String subj, String to, String cc, String msg)
 * adapted from https://ps06756.wordpress.com/2017/08/17/how-to-send-email-through-gmail-programmatically/</p>
 * @author Saieesh Rao, Pritzker School of Medicine '21
 * @version 1.4
 */
public class PritzkerEHR extends JFrame implements ActionListener, PropertyChangeListener{

    private static final String SMTP_HOST = "smtp.gmail.com";
    private static final String SMTP_PORT = "587";
    private static final String GMAIL_USERNAME = "pritzkerhosting@gmail.com";
    private static String GMAIL_PASSWORD;
    private SheetsHandler sheetsHandler;

    private JComboBox messageType;

    private JPanel buttonPanel;
    private JButton saveToDocButton;
    private JButton launchStepwiseButton;
    private JButton launchAllButton;
    private JTextField signatureField;

    private JTable jtR;
    private JTable jtP;
    private JTable jtI;
    private JTable jtD;
    private JProgressBar progressBar;
    private JTextArea intervieweeInfoArea;
    private JScrollPane intInfoScroll;

    private int nbrEmailsTotal;
    private int nbrEmailsDone;
    private int nbrWritesDone;
    private int nbrWritesTotal;
    private boolean suppressWriteProgress;

    private JPanel centerPanel;
    private JPanel requestReceivedPanel;
    private JPanel tellIntervieweesPanel;
    private JPanel ignoredRequestsPanel;
    private JPanel doneIntervieweePanel;
    private JPanel sigPanel;
    private JPanel askClassPanel;
    private JButton askClassButton;
    private JTextPane askClassMessage;
    private JTextField askClassSubjectField;
    private JTextField askClassRecipientsField;
    private JPanel askClassRecipientsPanel;
    private JPanel bottomPanel;

    private JDialog fluff;
    private Image windowIcon;
    private ImageIcon saveIcon;
    private ImageIcon stepForwardIcon;
    private ImageIcon fastForwardIcon;
    private ImageIcon sendMailIcon;

    protected static final String CONFIRM_INT_REQUEST_RECEIVED = "Confirm Interviewee's Request is Recieved";;  //confirm that we received the interviewee's host request through the google form
    protected static final String CONFIRM_HOST_TO_INTERVIEWEES = "Tell Interviewees their Host";                //confirm identity of host to interviewee
    protected static final String SEND_PLEA_TO_CLASS_LISTS = "Send Hosting Plea to Class Listhosts";            //email class listhosts with data on interviewees
    protected static final String VIEW_IGNORE_LIST = "View Removed Requests (for Duplicates/Cancellations)";    //duplicate requests and requests that have been cancelled
    protected static final String VIEW_COMPLETED_REQUESTS = "View Successfully Completed Hosting Requests";     //interviewees who have been successfully paired
    protected static final String HELP_AND_ABOUT = "Help and About";

    /**
     * <p>Constructor method for PritzkerEHR.</p>
     * <p>Implementation uses a subclass of javax.swing.SwingWorker to initialize the GUI and connect to Google Drive.
     * A JDialog containing a JProgressBar is instantiated concurrently while initialization takes place.</p>
     */
    public PritzkerEHR(){
        super("Pritzker Electronic Hosting Record - Gmail and Spreadsheet Manager");
        //load window icon
        try {
            windowIcon = ImageIO.read(getClass().getResource("resources/ehr_icon_fuzz.png"));
        } catch (IOException e) {
            e.printStackTrace();
        }
        setIconImage(windowIcon);
        JPanel jPanel = new JPanel(new GridLayout(2,1));
        JLabel jLabel = new JLabel("Password for pritzkerhosting@gmail.com?");
        JPasswordField jPasswordField = new JPasswordField();
        jPasswordField.requestFocusInWindow();
        jPanel.add(jLabel);
        jPanel.add(jPasswordField);
        jPanel.requestFocusInWindow();
        do {
            GMAIL_PASSWORD = null;
            JOptionPane pane = new JOptionPane(jPanel, JOptionPane.QUESTION_MESSAGE, JOptionPane.OK_CANCEL_OPTION, getWindowIcon(),new String[] {"OK","Cancel"},"OK"){
                @Override
                public void selectInitialValue() {
                    jPasswordField.requestFocusInWindow();
                }
            };
            pane.createDialog(this,"Pritzker EHR").setVisible(true);
            Object selectedButton = pane.getValue();
            //Object passwordInput = JOptionPane.showInputDialog(this, "Password for pritzkerhosting@gmail.com?", "Pritzker EHR", JOptionPane.QUESTION_MESSAGE, getWindowIcon(), null, null);
            if (selectedButton == null || selectedButton.equals("Cancel")) {
                System.exit(0);
            }
            String passwordInput = new String(jPasswordField.getPassword());
            GMAIL_PASSWORD = passwordInput.toString();
        }while(!confirmPassword());

        fluff = new JDialog(this,"Warming up Pritzker EHR",true);
        JProgressBar jProgressBar = new JProgressBar();
        jProgressBar.setIndeterminate(true);
        fluff.add(jProgressBar);
        fluff.pack();
        fluff.setLocationRelativeTo(null);

        SwingWorker startup = new StartupWorker(this);
        startup.addPropertyChangeListener(this);
        startup.execute();

        fluff.setVisible(true);

        addWindowListener(  new WindowAdapter(){
            public void windowClosing(WindowEvent we){
                System.exit(0);
            }
        });
        setResizable(false);
        pack();
        setLocationRelativeTo(null);
        setVisible(true);
    }

    /**
     * <p>Method for handling GUI ActionEvents. PritzkerEHR is the ActionListener for its GUI Components.</p>
     * @param ae an ActionEvent produced by a Component in the GUI.
     */
    public void actionPerformed(ActionEvent ae){
        Object source = ae.getSource();
        boolean launchAll = false;
        boolean isBeingEdited = jtR.isEditing() || jtP.isEditing() || jtI.isEditing();
        if(!isBeingEdited) {
            if (source.equals(launchStepwiseButton) || source.equals(launchAllButton)) {
                buttonPanel.removeAll();
                buttonPanel.add(progressBar);
                buttonPanel.revalidate();
                buttonPanel.repaint();
                if (source.equals(launchAllButton)) {
                    launchAll = true;
                }
                if (messageType.getSelectedItem().equals(CONFIRM_INT_REQUEST_RECEIVED)) {
                    suppressWriteProgress = true;
                    confirmRequestReceived(launchAll);

                } else if (messageType.getSelectedItem().equals(CONFIRM_HOST_TO_INTERVIEWEES)) {
                    suppressWriteProgress = true;
                    confirmHostToInterviewees(launchAll);
                }
            }
            else if (source.equals(saveToDocButton)) {
                suppressWriteProgress = false;
                resetProgressCounters();
                saveTablesToDoc();
            }
            else if(source.equals(askClassButton)){
                String recipients = askClassRecipientsField.getText();
                String body = askClassMessage.getText();
                String title = askClassSubjectField.getText();
                int choice = -1;
                choice = JOptionPane.showConfirmDialog(this,"All set to send the hosting plea \""+title+"\" to\n"+recipients+"?","Pritzker EHR says:",JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE);
                if(choice == JOptionPane.YES_OPTION) {
                    SwingWorker pleader = new EmailWorker(title, recipients, null, body, null); //sheetworker = null because no writing to sheet
                    pleader.addPropertyChangeListener(this);
                    pleader.execute();
                }
                else{
                    JOptionPane.showMessageDialog(this,"You decided not to send the hosting plea... yet","Pritzker EHR says:",JOptionPane.INFORMATION_MESSAGE);
                }
            }
            else if (source.equals(messageType)) {
                //if(askClassMessage.getText().trim().equals("")){    //plea somehow uninitialized because of threading issues
                SwingWorker pleaWriter = new PleaWorker(sheetsHandler.getUpcomingPairingArray());
                pleaWriter.execute();
                //}
                refreshCenterPanel();
            }
        }
        else{
            JOptionPane.showMessageDialog(this,"Please finish editing the cell before clicking any buttons","Pritzker EHR says:",JOptionPane.WARNING_MESSAGE);
        }
    }

    /**
     * <p>Method for updating the tables of interviewees and the hosting plea.</p>
     * <p>This method should be called after a batch of interviewees have had their hosting status changed,
     * either because an email was sent to them or because the hosting coordinator manually edited their
     * status by right-clicking them in the table and choosing an item in the pop-up menu.</p>
     */
    public void updatePleaAndTables(){
        JTable jtR_temp = sheetsHandler.getReceiptJTable();
        JTable jtP_temp = sheetsHandler.getPairingJTable();
        JTable jtI_temp = sheetsHandler.getIgnoredJTable();
        JTable jtD_temp = sheetsHandler.getDoneJTable();
        jtR.setModel(jtR_temp.getModel());
        sheetsHandler.conformThreeColumnTable(jtR);
        //jtR.setComponentPopupMenu(jtR_temp.getComponentPopupMenu());
        jtP.setModel(jtP_temp.getModel());
        sheetsHandler.conformFiveColumnTable(jtP);
        //jtP.setComponentPopupMenu(jtP_temp.getComponentPopupMenu());
        jtI.setModel(jtI_temp.getModel());
        sheetsHandler.conformThreeColumnTable(jtI);
        //jtI.setComponentPopupMenu(jtI_temp.getComponentPopupMenu());
        jtD.setModel(jtD_temp.getModel());
        sheetsHandler.conformFiveColumnTable(jtD);
        //jtD.setComponentPopupMenu(jtD_temp.getComponentPopupMenu());
        SwingWorker pleaWriter = new PleaWorker(sheetsHandler.getUpcomingPairingArray());
        pleaWriter.execute();
    }

    /**
     * <p>Method for handling progress updates from subclasses of javax.util.SwingWorker.</p>
     * <p>Subclasses of javax.swing.SwingWorker are primarily responsible for sending emails and writing updates to the Google Sheet.
     * This method is also responsible for closing the JDialog containing the JProgressBar during intiialization.</p>
     * @param pce a PropertyChangedEvent produced when an instance of javax.util.SwingWorker finishes its task.
     * */
    public void propertyChange(PropertyChangeEvent pce){
        Object source = pce.getSource();
        //System.out.println(source);
        if(source instanceof StartupWorker && ((StartupWorker)source).getProgress() == 100){
            //System.out.println("pce StartupWorker");
            fluff.dispatchEvent(new WindowEvent(fluff, WindowEvent.WINDOW_CLOSING));
        }
        else if(source instanceof EmailWorker){
            refreshProgressBar();
//            if(((EmailWorker)source).getProgress() == 100){
//                JOptionPane.showMessageDialog(this, "All requested emails sent!", "Pritzker EHR says:", JOptionPane.INFORMATION_MESSAGE,getWindowIcon());
//            }
        }
        else if(source instanceof SheetWorker){
            refreshProgressBar();
//            if(((SheetWorker)source).getProgress() == 100){
//                JOptionPane.showMessageDialog(this, "All requested emails sent!", "Pritzker EHR says:", JOptionPane.INFORMATION_MESSAGE,getWindowIcon());
//            }
        }
    }

    /**
     * Gets an instance of the icon displayed in the top-left corner of the PritzkerEHR window.
     * @return the window icon for PritzkerEHR
     */
    public ImageIcon getWindowIcon(){
        return new ImageIcon(windowIcon);
    }
    /**
     * <p>A method for safely sending emails to interviewees.</p>
     * <p>The content of the email is read from Google Sheet PritzkerHosting_PATH, Sheet 1, cell B3, specifying that the interviewee's request for a host has been received.
     * The interviewees who may receive this email are those who are shown in the "Confirm Interviewee's Request is Received" table, which are those who
     * have the number 0 in Sheet 2, Column M of the Google Sheet specified by the spreadsheetID in PritzkerHosting_PATH, Sheet 1, cell B1.
     * When an email is sent using this method, the Google Sheet containing interviewee information is updated to reflect the change in hosting status. Specifically,
     * the 0 in Sheet2, Column M is changed to a 1.</p>
     *
     * <p>WARNING: Writes to the Google Sheet performed by this method are not guaranteed safe if the sheet is being edited concurrently by the user. User operations on the
     * sheet that rearrange the order of the sheet's rows while PritzkerEHR is open may result in incorrect writes and deleterious overwriting
     * of important interviewee data in the sheet.</p>
     *
     * @param launchAll whether the emails should be sent with individual review (false) or sent all at once without individual review (true)
     */
    private void confirmRequestReceived(boolean launchAll){
        resetProgressCounters();
        saveTablesToDoc();
        String[] namesList = new String[sheetsHandler.numberForReceipt()];
        String[] datesList = new String[sheetsHandler.numberForReceipt()];
        String[] emailList = new String[sheetsHandler.numberForReceipt()];
        String[] rowList = new String[sheetsHandler.numberForReceipt()];
        for (int i = 0; i < sheetsHandler.numberForReceipt(); i++) {
            datesList[i] = (String) jtR.getValueAt(i, 0);
            namesList[i] = (String) jtR.getValueAt(i, 1);
            emailList[i] = (String) jtR.getValueAt(i, 2);
            rowList[i] = jtR.getValueAt(i, 3).toString();
            //System.out.println(emailList[jtR.convertRowIndexToModel(i)]);
        }

        int choice = JOptionPane.CANCEL_OPTION;
        for (int i = 0; i < datesList.length; i++) {
            String msg = sheetsHandler.getReceiptEmailText();
            msg = msg.replaceAll("\\[INTERVIEWEE NAME\\]",getFirstName(namesList[i]));//namesList[i].split(" ")[0]
            msg = msg.replaceAll("\\[HOSTING DATE\\]",datesList[i]);
            msg = msg.replaceAll("\\[PREFERENCE TABLE\\]",sheetsHandler.getHTMLPreferenceTable(sheetsHandler.getUpcomingReceiptArray(),i));
            msg = msg.replaceAll("\\[SIGNATURE\\]",signatureField.getText());
            JTextPane msgPane = new JTextPane();
            //msgPane.setPreferredSize(new Dimension(200, 300));
            msgPane.setContentType("text/html");

            msgPane.setText(msg);
            JScrollPane scrollPane = new JScrollPane(msgPane);
            scrollPane.setPreferredSize(new Dimension(200, 300));
            JLabel toLabel = new JLabel("To:");
            JTextField toField = new JTextField(30);
            toField.setText(emailList[i]);
            JPanel toPanel = new JPanel();
            JPanel dispPanel = new JPanel(new BorderLayout());
            toPanel.add(toLabel);
            toPanel.add(toField);
            dispPanel.add(toPanel, BorderLayout.NORTH);
            dispPanel.add(scrollPane, BorderLayout.SOUTH);
            if (!launchAll) {
                choice = JOptionPane.showConfirmDialog(this, dispPanel, "Subject: Pritzker Hosting Confirmation", JOptionPane.OK_CANCEL_OPTION);
                if (choice == JOptionPane.OK_OPTION) {
                    emailList[i] = toField.getText();
                    msg = msgPane.getText();
                    SwingWorker sheetWorker = new SheetWorker("Sheet2!M" + rowList[i], "1", sheetsHandler);
                    sheetWorker.addPropertyChangeListener(this);
                    SwingWorker emailWorker = new EmailWorker("Pritzker Hosting Confirmation", emailList[i], null, msg, (SheetWorker)sheetWorker);
                    emailWorker.addPropertyChangeListener(this);
                    emailWorker.execute();
                    buttonPanel.revalidate();
                    buttonPanel.repaint();
                } else {
                    nbrEmailsTotal--;
                    nbrWritesTotal--;
                    refreshProgressBar();
                    buttonPanel.revalidate();
                    buttonPanel.repaint();
                    Object[] options = {"OK", "Skip All"};
                    int n = JOptionPane.showOptionDialog(this,
                            "Email to " + namesList[i] + " skipped",
                            "Pritzker EHR says:",
                            JOptionPane.YES_NO_OPTION,
                            JOptionPane.WARNING_MESSAGE,
                            getWindowIcon(),     //use custom icon
                            options,  //the titles of buttons
                            options[0]); //default button title
                    if (n == 1) {
                        refreshBottomPanel();
                        break;
                    }
                    JOptionPane.showMessageDialog(this, "Email to " + namesList[i] + " skipped", "Pritzker EHR says:", JOptionPane.INFORMATION_MESSAGE,getWindowIcon());
                }
            } else {
                msg = msgPane.getText();
                SwingWorker sheetWorker = new SheetWorker("Sheet2!M" + rowList[i], "1", sheetsHandler);
                sheetWorker.addPropertyChangeListener(this);
                SwingWorker emailWorker = new EmailWorker("Pritzker Hosting Confirmation", emailList[i], null, msg, (SheetWorker)sheetWorker);
                emailWorker.addPropertyChangeListener(this);
                emailWorker.execute();
                //buttonPanel.revalidate();
                //buttonPanel.repaint();
            }
        }
        sheetsHandler.refresh();
        requestReceivedPanel.removeAll();
        JScrollPane jspR = new JScrollPane(jtR, JScrollPane.VERTICAL_SCROLLBAR_ALWAYS, JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
        requestReceivedPanel.add(jspR);
//        if (emailList.length > 0) { //&& (choice == JOptionPane.OK_OPTION || launchAll)
//            JOptionPane.showMessageDialog(this, "All requested emails sent!", "Pritzker Hosting Bot says:", JOptionPane.INFORMATION_MESSAGE);
//        }
        refreshBottomPanel();
        pack();
    }

    /**
     * <p>A method for safely sending emails to interviewees.</p>
     * <p>The content of the email is read from Google Sheet PritzkerHosting_PATH, Sheet 1, cell B4, which specifies the interviewee's host and instructs
     * the interviewee to reach out to the host. The host is cc'd on the email.
     * The interviewees who may receive this email are those who are shown in the "Tell Interviewee's their Host" table, which are those who
     * have the number 1 in Sheet 2, Column M of the Google Sheet specified by the spreadsheetID in PritzkerHosting_PATH, Sheet 1, cell B1.
     * When an email is sent using this method, the Google Sheet containing interviewee information is updated to reflect the change in hosting status. Specifically,
     * the 1 in Sheet 2, Column M is changed to a 2.</p>
     *
     * <p>WARNING: Writes to the Google Sheet performed by this method are not guaranteed safe if the sheet is being edited concurrently by the user. User operations on the
     * sheet that rearrange the order of the sheet's rows while PritzkerEHR is open may result in incorrect writes and deleterious overwriting
     * of important interviewee data in the sheet.</p>
     *
     * @param launchAll whether the emails should be sent with individual review (false) or sent all at once without individual review (true)
     */
    private void confirmHostToInterviewees(boolean launchAll){
        resetProgressCounters();
        saveTablesToDoc();
        String[] intNamesList = new String[sheetsHandler.numberForPairing()];
        String[] datesList = new String[sheetsHandler.numberForPairing()];
        String[] intEmailList = new String[sheetsHandler.numberForPairing()];
        String[] hostNamesList = new String[sheetsHandler.numberForPairing()];
        String[] hostEmailList = new String[sheetsHandler.numberForPairing()];
        String[] rowList = new String[sheetsHandler.numberForPairing()];
        for (int i = 0; i < sheetsHandler.numberForPairing(); i++) {
            hostNamesList[i] = (String) jtP.getValueAt(i, 0);
            hostEmailList[i] = (String) jtP.getValueAt(i, 1);
            datesList[i] = (String) jtP.getValueAt(i, 2);
            intNamesList[i] = (String) jtP.getValueAt(i, 3);
            intEmailList[i] = (String) jtP.getValueAt(i, 4);
            rowList[i] = jtP.getValueAt(i, 5).toString();
            //System.out.printf("%s\t%s\t%s\t%s\t%s\n", hostNamesList[i], hostEmailList[i], datesList[i], intNamesList[i], intEmailList[i]);
        }
        String[][] net = new String[][]{intNamesList, datesList, intEmailList, hostNamesList, hostEmailList};
        boolean proceed = true;
        ArrayList<Integer> viableIndices = new ArrayList<Integer>();
        for (int i = 0; i < sheetsHandler.numberForPairing(); i++) {
            boolean completeRow = true;
            for (int j = 0; j < 5; j++) {
                //System.out.println(i + "," + j);
                if (net[j][i] == null || net[j][i].equals("")) {
                    completeRow = false;
                    break;
                }
            }
            if (completeRow) {
                viableIndices.add(i);
            }
        }
        //System.out.println("VC:" + viableIndices.size());
        if (proceed && viableIndices.size() > 0) {
            int choice = JOptionPane.CANCEL_OPTION;
            double percentComplete = 0;
            for (int i = 0; i < viableIndices.size(); i++) {
                int k = viableIndices.get(i);
                String msg = sheetsHandler.getPairingEmailText();
                msg = msg.replaceAll("\\[INTERVIEWEE NAME\\]",getFirstName(intNamesList[k]));//intNamesList[k].split(" ")[0]
                msg = msg.replaceAll("\\[HOSTING DATE\\]",datesList[k]);
                msg = msg.replaceAll("\\[HOST NAME\\]",hostNamesList[k]);
                msg = msg.replaceAll("\\[SIGNATURE\\]",signatureField.getText());
                JTextPane msgPane = new JTextPane();
                msgPane.setPreferredSize(new Dimension(200, 500));
                msgPane.setContentType("text/html");

                msgPane.setText(msg);
                JLabel toLabel = new JLabel("To:");
                JLabel ccLabel = new JLabel("Cc:");
                JTextField toField = new JTextField(30);
                toField.setText(intEmailList[k]);
                JTextField ccField = new JTextField(30);
                ccField.setText(hostEmailList[k]);
                JPanel toPanel = new JPanel();
                JPanel ccPanel = new JPanel();
                JPanel dispPanel = new JPanel(new BorderLayout());
                toPanel.add(toLabel);
                toPanel.add(toField);
                ccPanel.add(ccLabel);
                ccPanel.add(ccField);
                dispPanel.add(toPanel, BorderLayout.NORTH);
                dispPanel.add(ccPanel, BorderLayout.CENTER);
                dispPanel.add(msgPane, BorderLayout.SOUTH);
                if (!launchAll) {
                    choice = JOptionPane.showConfirmDialog(this, dispPanel, "Subject: Pritzker Hosting Info", JOptionPane.OK_CANCEL_OPTION);
                    if (choice == JOptionPane.OK_OPTION) {
                        intEmailList[k] = toField.getText();
                        hostEmailList[k] = ccField.getText();
                        msg = msgPane.getText();
                        SwingWorker sheetWorker = new SheetWorker("Sheet2!M" + rowList[k], "2", sheetsHandler);
                        sheetWorker.addPropertyChangeListener(this);
                        SwingWorker emailWorker = new EmailWorker("Pritzker Hosting Info", intEmailList[k], hostEmailList[k], msg, (SheetWorker)sheetWorker);
                        emailWorker.addPropertyChangeListener(this);
                        emailWorker.execute();
                        buttonPanel.repaint();
                    } else {
                        nbrEmailsTotal--;
                        nbrWritesTotal--;
                        refreshProgressBar();
                        buttonPanel.repaint();
                        Object[] options = {"OK", "Skip All"};
                        int n = JOptionPane.showOptionDialog(this,
                                "Email to " + intNamesList[k] + " skipped",
                                "Pritzker EHR says:",
                                JOptionPane.YES_NO_OPTION,
                                JOptionPane.WARNING_MESSAGE,
                                getWindowIcon(),     //the custom icon
                                options,  //the titles of buttons
                                options[0]); //default button title
                        if (n == 1) {
                            refreshBottomPanel();
                            break;
                        }
                    }
                }
                else {
                    intEmailList[k] = toField.getText();
                    hostEmailList[k] = ccField.getText();
                    msg = msgPane.getText();
                    SwingWorker sheetWorker = new SheetWorker("Sheet2!M" + rowList[k], "2", sheetsHandler);
                    sheetWorker.addPropertyChangeListener(this);
                    SwingWorker emailWorker = new EmailWorker("Pritzker Hosting Info", intEmailList[k], hostEmailList[k], msg, (SheetWorker)sheetWorker);
                    emailWorker.addPropertyChangeListener(this);
                    emailWorker.execute();
                    //buttonPanel.repaint();
                }
            }
            sheetsHandler.refresh();
            tellIntervieweesPanel.removeAll();
            JScrollPane jspP = new JScrollPane(jtP, JScrollPane.VERTICAL_SCROLLBAR_ALWAYS, JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
            //jspP.setVisible(true);
            tellIntervieweesPanel.add(jspP);
//            if (datesList.length > 0) { //&& choice == JOptionPane.OK_OPTION || launchAll
//                JOptionPane.showMessageDialog(this, "All requested emails sent!", "Pritzker Hosting Bot says:", JOptionPane.INFORMATION_MESSAGE);
//            }
            refreshBottomPanel();
            pack();
        }
    }

    /**
     * <p>Writes the contents of all edited cells in all interviewee tables to the Google Sheet.
     * The cells are determined to be edited if their contents differ from the contents of the cells when they were last read from the Google Sheet.</p>
     *
     * <p>WARNING: Writes to the Google Sheet performed by this method are not guaranteed safe if the sheet is being edited concurrently by the user. User operations on the
     * sheet that rearrange the order of the sheet's rows while PritzkerEHR is open may result in incorrect writes and deleterious overwriting
     * of important interviewee data in the sheet.</p>
     */
    private void saveTablesToDoc(){
        //System.out.println("saving to doc button clicked");
        buttonPanel.removeAll();
        buttonPanel.add(progressBar);
        buttonPanel.revalidate();
        buttonPanel.repaint();
        //requests received
        for (int i = 0; i < jtR.getRowCount(); i++) {
            for (int j = 0; j < jtR.getColumnCount()-1; j++) {  //columnCount -1 because the last column is the row # (not recorded in a column on the Google Sheet)
                Object tableObject = jtR.getModel().getValueAt(i, j);
                Object dataObject = sheetsHandler.getReceiptTableValueFromSheet(i, j);
                if (!tableObject.equals(dataObject)) {
                    ///nbrWritesTotal++;
                    //System.out.println("[" + i + "," + j + "]: " + tableObject + " != " + dataObject);
                    String range = sheetsHandler.getReceiptTableValueSheetRange(i, j);
                    //System.out.println("edit: " + range);
                    SwingWorker sheetEditor = new SheetWorker(range, tableObject.toString(), sheetsHandler);
                    sheetEditor.addPropertyChangeListener(this);
                    sheetEditor.execute();
                }
            }
        }
        //pairing stage
        for (int i = 0; i < jtP.getRowCount(); i++) {
            for (int j = 0; j < jtP.getColumnCount()-1; j++) {  //columnCount -1 because the last column is the row # (not recorded in a column on the Google Sheet)
                Object tableObject = jtP.getModel().getValueAt(i, j);
                Object dataObject = sheetsHandler.getPairingTableValueFromSheet(i, j);
                if (!tableObject.equals(dataObject)) {
                    ///nbrWritesTotal++;
                    //System.out.println("[" + i + "," + j + "]: " + tableObject + " != " + dataObject);
                    String range = sheetsHandler.getPairingTableValueSheetRange(i, j);
                    //System.out.println("edit: " + range);
                    SwingWorker sheetEditor = new SheetWorker(range, tableObject.toString(), sheetsHandler);
                    sheetEditor.addPropertyChangeListener(this);
                    sheetEditor.execute();
                }
            }
        }
        //ignored requests
        for (int i = 0; i < jtI.getRowCount(); i++) {
            for (int j = 0; j < jtI.getColumnCount()-1; j++) {  //columnCount -1 because the last column is the row # (not recorded in a column on the Google Sheet)
                Object tableObject = jtI.getModel().getValueAt(i, j);
                Object dataObject = sheetsHandler.getIgnoredTableValueFromSheet(i, j);
                if (!tableObject.equals(dataObject)) {
                    ///nbrWritesTotal++;
                    //System.out.println("[" + i + "," + j + "]: " + tableObject + " != " + dataObject);
                    String range = sheetsHandler.getIgnoredTableValueSheetRange(i, j);
                    //System.out.println("edit: " + range);
                    SwingWorker sheetEditor = new SheetWorker(range, tableObject.toString(), sheetsHandler);
                    sheetEditor.addPropertyChangeListener(this);
                    sheetEditor.execute();
                }
            }
        }
        //done pairings
        for (int i = 0; i < jtD.getRowCount(); i++) {
            for (int j = 0; j < jtD.getColumnCount()-1; j++) {  //columnCount -1 because the last column is the row # (not recorded in a column on the Google Sheet)
                Object tableObject = jtD.getModel().getValueAt(i, j);
                Object dataObject = sheetsHandler.getDoneTableValueFromSheet(i, j);
                if (!tableObject.equals(dataObject)) {
                    ///nbrWritesTotal++;
                    //System.out.println("[" + i + "," + j + "]: " + tableObject + " != " + dataObject);
                    String range = sheetsHandler.getDoneTableValueSheetRange(i, j);
                    //System.out.println("edit: " + range);
                    SwingWorker sheetEditor = new SheetWorker(range, tableObject.toString(), sheetsHandler);
                    sheetEditor.addPropertyChangeListener(this);
                    sheetEditor.execute();
                }
            }
        }
        sheetsHandler.refresh();
        //System.out.println("save to doc code finished");
        //refreshBottomPanel();
    }

    /**
     * <p>Refreshes the <code>JProgressBar</code> at the bottom of the window.</p>
     * <p>If only writes to the Google Sheet are being performed, then the progress is calculated as the percent of <code>SheetWorker</code>s
     * that have been completed in the current writing operation. If at least one email is being sent, then the progress is calculated as 0.33
     * times the percentage of <code>SheetWorkers</code>s that have completed in the current operation plus 0.67 times the percentage of
     * <code>EmailWorker</code>s that have completed in the current operation.</p>
     *
     * <p>This method is often called from the <code>done()</code> method of subclasses of <code>javax.swing.SwingWorker</code>, hence the method is <code>synchronized</code>.</p>
     *
     */
    private synchronized void refreshProgressBar(){
        //all emails require a paired write to the Google Doc
        //not all writes to the Google doc are accompanied by an email (i.e. "Save Edits to Doc" button)
        int progress = 0;
        if (nbrEmailsTotal == 0 && nbrWritesTotal > 0 && !suppressWriteProgress) {
            progress = (int)Math.round(100.0*nbrWritesDone/nbrWritesTotal);
        } else if (nbrEmailsTotal > 0 && nbrWritesTotal > 0) {
            double emailContribution = 67.0, writeContribution = 33.0;
            double emailProgress = emailContribution * nbrEmailsDone / nbrEmailsTotal;
            double sheetProgress = writeContribution * nbrWritesDone / nbrWritesTotal;
            progress = (int)Math.round(emailProgress + sheetProgress);
        }
        progressBar.setValue(Math.min(progress, 100));
    }

//    /**
//     * <p>Ensures that the contents of the bottom panel are in agreement with the current operations being performed by PritzkerEHR.</p>
//     *
//     * <p>When the tables for confirming receipt of request and pairing are first opened, this method ensures that the bottom panel shows the buttons
//     * for saving tables to the Google Sheet, sending emails with individual review, and sending all emails at once are shown, as well as the textfield
//     * for entering the email sender's (user's) name. Once a button is clicked, the buttons disappear and are replaced by a progress bar indicating the
//     * progress of the current operation.</p>
//     *
//     * <p>For more information on when each button is displayed, please see the documentation for <code>refreshButtonPanel()</code></p>.
//     */

    /**
     * <p>Resets the contents of the bottom panel to show the signature panel and the button panel. This method should be called to
     * safely remove the progress bar from the bottom panel.</p>
     */
    private void refreshBottomPanel(){
        bottomPanel.removeAll();
        bottomPanel.add(sigPanel, BorderLayout.NORTH);
        bottomPanel.add(buttonPanel, BorderLayout.SOUTH);
        bottomPanel.revalidate();
        bottomPanel.repaint();
        pack();
    }

    /**
     * <p>Resets the button panel to contain the three standard buttons for 1) saving edits,
     * 2) launching emails with individual review, and 3) launching emails without individual review.</p>
     */
    private void refreshButtonPanel(){
        buttonPanel.removeAll();
        buttonPanel.add(saveToDocButton);
        buttonPanel.add(launchStepwiseButton);
        buttonPanel.add(launchAllButton);
    }

    /**
     * <p>Ensures that the center panel displays the appropriate table and the detail text inset.</p>
     * <p>As a safety feature, this method also calls <code>refreshButtonPanel()</code> and <code>refreshBottomPanel()</code>
     * so that the contents of the bottom panel are always appropriate for the contents of the center panel.</p>
     */
    private void refreshCenterPanel(){
        if (messageType.getSelectedItem().equals(CONFIRM_INT_REQUEST_RECEIVED)) {
            centerPanel.removeAll();
            centerPanel.add(requestReceivedPanel);
            centerPanel.add(intInfoScroll);
            refreshButtonPanel();
            refreshBottomPanel();
        } else if (messageType.getSelectedItem().equals(CONFIRM_HOST_TO_INTERVIEWEES)) {
            centerPanel.removeAll();
            centerPanel.add(tellIntervieweesPanel);
            centerPanel.add(intInfoScroll);
            refreshButtonPanel();
            refreshBottomPanel();
        } else if (messageType.getSelectedItem().equals(VIEW_COMPLETED_REQUESTS)) {
            centerPanel.removeAll();
            centerPanel.add(doneIntervieweePanel);
            centerPanel.add(intInfoScroll);
            buttonPanel.removeAll();
            buttonPanel.add(saveToDocButton);
            refreshBottomPanel();
            bottomPanel.remove(sigPanel);
        } else if (messageType.getSelectedItem().equals(VIEW_IGNORE_LIST)) {
            centerPanel.removeAll();
            centerPanel.add(ignoredRequestsPanel);
            centerPanel.add(intInfoScroll);
            buttonPanel.removeAll();
            buttonPanel.add(saveToDocButton);
            refreshBottomPanel();
            bottomPanel.remove(sigPanel);
        }
        else if (messageType.getSelectedItem().equals(SEND_PLEA_TO_CLASS_LISTS)) {
            centerPanel.removeAll();
            centerPanel.add(askClassPanel);
            bottomPanel.removeAll();
            bottomPanel.add(askClassRecipientsPanel, BorderLayout.NORTH);
            bottomPanel.add(askClassButton, BorderLayout.SOUTH);
        }
        else if (messageType.getSelectedItem().equals(HELP_AND_ABOUT)){
            centerPanel.removeAll();
            bottomPanel.removeAll();
            JTextPane jtp = new JTextPane();
            jtp.setContentType("text/html");
            try {
                jtp.setPage(getClass().getResource("resources/about.html"));
            } catch (IOException e) {
                e.printStackTrace();
            }
            jtp.setEditable(false);
            jtp.setCaretPosition(0);
            JScrollPane jsp = new JScrollPane(jtp,JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
            jsp.setPreferredSize(new Dimension(692,293));
            centerPanel.add(jsp);
        }
        intervieweeInfoArea.setText("");
        centerPanel.revalidate();
        centerPanel.repaint();
        pack();
    }

    /**
     * <p>Increment <code>nbrEmailsTotal</code> by one.
     * This method is <code>synchronized</code>because it is called by instances of <code>EmailWorker</code>.</p>
     */
    private synchronized void incrementEmailsTotal(){
        nbrEmailsTotal++;
    }

    /**
     * <p>Increment <code>nbrEmailsDone</code> by one.
     * This method is <code>synchronized</code>because it is called by instances of <code>EmailWorker</code>.</p>
     */
    private synchronized void incrementEmailsDone(){
        nbrEmailsDone++;
    }

    /**
     * <p>Increment <code>nbrWritesTotal</code> by one.
     * This method is <code>synchronized</code>because it is called by instances of <code>SheetWorker</code>.</p>
     */
    private synchronized void incrementWritesTotal(){
        nbrWritesTotal++;
    }

    /**
     * <p>Increment <code>nbrWritesDone</code> by one.
     * This method is <code>synchronized</code>because it is called by instances of <code>SheetWorker</code>.</p>
     */
    private synchronized void incrementWritesDone(){
        nbrWritesDone++;
    }

    /**
     * <p>Sets the values of <code>nbrEmailsTotal</code>, <code>nbrEmailsDone</code>, <code>nbrWritesTotal</code>,
     * and <code>nbrWritesDone</code> to 0. This method is synchronized so as not to conflict with the relevant increment methods.</p>
     */
    private synchronized void resetProgressCounters(){
        nbrEmailsTotal = 0;
        nbrEmailsDone = 0;
        nbrWritesTotal = 0;
        nbrWritesDone = 0;
    }

    /**
     * <p>Gets an instance representation of the Gmail login session.</p>
     * @return an instance of the Gmail login session.
     */
    private Session getSession(){
        Properties prop = System.getProperties();
        prop.setProperty("mail.smtp.starttls.enable", "true");
        prop.setProperty("mail.smtp.host", SMTP_HOST);
        prop.setProperty("mail.smtp.user", GMAIL_USERNAME);
        prop.setProperty("mail.smtp.password", GMAIL_PASSWORD);
        prop.setProperty("mail.smtp.port", SMTP_PORT);
        prop.setProperty("mail.smtp.auth", "true");

        return Session.getInstance(prop, new Authenticator() {
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(GMAIL_USERNAME, GMAIL_PASSWORD);
            }
        });
    }

    /**
     * <p>Connects the current login session to the Gmail server.</p>
     * <p>If <code>GMAIL_PASSWORD</code> is incorrect, a <code>JOptionPane</code> is opened telling the user that
     * the password is incorrect, and a <code>false</code> value is returned. This method may also
     * return <code>false</code> if the connection to the server times out.</p>
     *
     * @return whether a successful connection to the server was made using the given password
     */
    private boolean confirmPassword(){
        Session s = getSession();
        try {
            Transport transport =s.getTransport("smtp");
            transport.connect(SMTP_HOST, GMAIL_USERNAME, GMAIL_PASSWORD);
        }
        catch(Exception e){
            String msg = "The password was incorrect for pritzkerhosting@gmail.com,\nyour device is disconnected from the network, or the network\nconnection timed out. Please try again.";
            JOptionPane.showMessageDialog(this,msg,"Pritzker EHR",JOptionPane.ERROR_MESSAGE,getWindowIcon());
            return false;
        }
        return true;
    }

    /**
     * <p>Sends an email from pritzkerhosting@gmail.com to the intended recipient(s).</p>
     * <p>WARNING: This method should not be executed on the GUI thread as it can cause the GUI to lock
     * while the email is being sent. Emails should be sent using an instance of <code>EmailWorker</code>
     * which is implemented using this method.</p>
     * @param subj The subject line of the email
     * @param to The intended recipient(s) of the email, as specified in the email's "To:" field.
     *          If more than one recipient is intended, their email addresses should be separated by commas.
     * @param cc The intended recipient(s) of the email, as specified in the email's "Cc:" field.
     *          If more than one recipient is intended, their email addresses should be separated by commas.
     * @param msg The email body in HTML format.
     * @return a boolean representing whether the email was successfully sent (true) or not (false)
     */
    private boolean sendMail(String subj, String to, String cc, String msg) {
        Session session = getSession();
        MimeMessage message = new MimeMessage(session);
        try {
            message.setFrom(new InternetAddress(GMAIL_USERNAME));
            message.addRecipients(Message.RecipientType.TO,
                    InternetAddress.parse(GMAIL_USERNAME));
            message.setSubject(subj);
            MimeBodyPart htmlPart = new MimeBodyPart();
            Multipart multiPart = new MimeMultipart("alternative");
            htmlPart.setContent(msg, "text/html; charset=utf-8");
            multiPart.addBodyPart(htmlPart);
            message.setContent(multiPart);
            message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(to));
            if (cc != null) {
                message.setRecipients(Message.RecipientType.CC, InternetAddress.parse(cc));
            }
            Transport transport = session.getTransport("smtp");
            transport.connect(SMTP_HOST, GMAIL_USERNAME, GMAIL_PASSWORD);
            transport.sendMessage(message, message.getAllRecipients());
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, e.toString(), "Error: Stack Trace (email text next)", JOptionPane.ERROR_MESSAGE, getWindowIcon());
            JOptionPane.showMessageDialog(null, msg, "Error: Problem Email Text", JOptionPane.ERROR_MESSAGE, getWindowIcon());
            return false;
        }
        return true;
    }

    /**
     * <p>Gets the <code>String</code> representation of the Pritzker class listhosts to which the hosting plea
     * emails should be addressed.</p>
     * <p>The three listhosts addresses in the <code>String</code> are separated by commas and hence can be used
     * directly as the "To:" or "Cc:" field in the <code>sendMail(...)</code> method. The listhosts are chosen based
     * on the current date on the computer such that the addressed listhosts represent the 1st, 2nd, and 4th year
     * classes. The current date is retrieved using an instance of the <code>java.util.Calendar</code> class.</p>
     * @return a <code>String</code> representing the three class listhosts that will receive the hosting plea
     */
    private String initializePleaRecipients(){
        Calendar today = Calendar.getInstance();
        int month = today.get(Calendar.MONTH);
        int year = today.get(Calendar.YEAR)-2000;
        if(month > Calendar.JUNE){  //starting the new school year, so...
            year++;                 //M4s are graduating in June of the following year
        }
        StringBuilder sb = new StringBuilder(100);
        String listhostBase = "@lists.uchicago.edu";
        sb.append("ms");
        sb.append(year);            //M4s
        sb.append(listhostBase);
        sb.append(", ");
        sb.append("ms");
        sb.append(year+2);          //M2s
        sb.append(listhostBase);
        sb.append(", ");
        sb.append("ms");
        sb.append(year+3);          //M1s
        sb.append(listhostBase);
        return sb.toString();
    }

    /**
     * <p>Gets the first name of the interviewee given the interviewee's full name. For most applicants, their first
     * name will be the first token in the full name i.e. the word that comes before the first space character.</p>
     * <p>Often, interviewees will include both their legal names and preferred name in the same entry, usually denoting
     * their preferred name within parentheses or quotation marks. If there are a pair of open/close parentheses or
     * quotation marks in the given name, the method will preferentially return that as their first name.</p>
     * @return the first name of the interviewee
     */
    private String getFirstName(String fullName){
        //check if the name contains (, ), or "
        int opInd = fullName.indexOf("(");
        int cpInd = fullName.indexOf(")");
        int oqInd = fullName.indexOf("\"");
        int cqInd = fullName.indexOf("\"",oqInd+1);
        //preferred name between parentheses
        if(opInd >= 0 && cpInd > 0 && cpInd > opInd){
            return fullName.substring(opInd+1,cpInd);
        }
        //preferred name between quotes
        else if(oqInd >= 0 && cqInd > 0 && cqInd > oqInd){
            return fullName.substring(oqInd+1,cqInd);
        }
        //default: the first token in the full name before the first space character
        return fullName.split(" ")[0];
    }
    /**
     * Starts the PritzkerEHR application.
     * @param args n/a
     */
    public static void main(String [] args){
        PritzkerEHR phh2 = new PritzkerEHR();
    }

    /**
     * <p>A subclass of <code>javax.swing.SwingWorker</code> to send emails off the GUI thread.</p>
     * <p>Instances of <code>EmailWorker</code> can only send an email with the parameters specified at construction
     * and only execute once. Hence, multiple emails must be sent individually by multiple <code>EmailWorker</code>s.</p>
     */
    private class EmailWorker extends SwingWorker<Boolean, Void> {
        private String subj;
        private String to;
        private String cc;
        private String msg;
        private SheetWorker sw;

        /**
         * <p>Constructor for an instance of <code>EmailWorker</code>. The variable <code>nbrEmailsTotal</code>
         * is incremented at construction.</p>
         * @param subj The subject line of the email
         * @param to The intended recipient(s) of the email, as specified in the email's "To:" field.
         *          If more than one recipient is intended, their email addresses should be separated by commas.
         * @param cc The intended recipient(s) of the email, as specified in the email's "Cc:" field.
         *          If more than one recipient is intended, their email addresses should be separated by commas.
         * @param msg The email body in HTML format.
         */
        public EmailWorker(String subj, String to, String cc, String msg, SheetWorker sw){
            this.subj = subj;
            this.to = to;
            this.cc = cc;
            this.msg = msg;
            this.sw = sw;
            incrementEmailsTotal();
        }

        /**
         * <p>Sends the email using <code>sendEmail(...)</code>.</p>
         * @return a boolean representing whether the email was successfully sent (true) or not (false)
         */
        public Boolean doInBackground(){
            //Thread.sleep((long)(Math.random()*5000));   //added to prevent overloading of email server, leading to crashes
            boolean worked = sendMail(subj, to, cc, msg);
            if(sw != null && worked){
                sw.execute();
            }
            return worked;
        }

        /**
         * <p>Increments <code>nbrEmailsDone</code> and sets the progress attribute to 100, thus firing a <code>PropertyChangedEvent</code>.</p>
         */
        public void done(){
            incrementEmailsDone();
            setProgress(100);
        }
    }

    /**
     * <p>A subclass of <code>javax.swing.SwingWorker</code> to write updates to the Google Sheet off the GUI thread.</p>
     * <p>Instances of <code>SheetWorker</code> can only update the sheet with the parameters specified at construction
     * and only execute once. Hence, multiple updates must be sent individually by multiple <code>SheetWorker</code>s.</p>
     */
    private class SheetWorker extends SwingWorker<Boolean, Void> {
        private String range;
        private String input;
        private SheetsHandler sheetsHandler;

        /**
         * <p>Constructor for an instance of <code>SheetWorker</code>. The variable <code>nbrWritesTotal</code>
         * is incremented at construction.</p>
         * @param range The cells in the sheet that will be written to
         * @param input The contents that will be written to the cells in <code>range</code>
         * @param sheetsHandler The instance of <code>SheetsHandler</code> that will be used to write to the sheet
         */
        public SheetWorker(String range, String input, SheetsHandler sheetsHandler){
            this.range = range;
            this.input = input;
            this.sheetsHandler = sheetsHandler;
            incrementWritesTotal();
        }

        /**
         * <p>Writes to the sheet using the given <code>SheetsHandler</code>.</p>
         * @return whether the write to the sheet was successful (true) or not (false)
         */
        public Boolean doInBackground(){
            try {
                sheetsHandler.write(range, input);
            }
            catch(Exception e){
                e.printStackTrace();
                return false;
            }
            return true;
        }

        /**
         * <p>Increments <code>nbrWritesDone</code> and sets the progress attribute to 100, thus firing a <code>PropertyChangedEvent</code>.</p>
         */
        public void done(){
            incrementWritesDone();
            setProgress(100);
        }
    }

    /**
     * <p>A subclass of <code>javax.swing.SwingWorker</code> to draft the plea email off the GUI thread.</p>
     * <p>The plea email is drafted using data on the interviewees that were previously downloaded from the
     * Google Sheet. To ensure the accuracy of the plea email, the interviewee data should be downloaded from
     * the sheet prior to instantiating the <code>PleaWorker</code>.</p>
     */
    private class PleaWorker extends SwingWorker<String, Void> {
        private Object[][] upcoming;
        private ArrayList<Integer> forTable;
        private StringBuilder sb;

        /**
         * <p>Constructor for an instance of <code>PleaWorker</code>.</p>
         * @param upcoming an array containing all the interviewee data downloaded from the Google Sheet
         */
        public PleaWorker(Object[][] upcoming){
            this.upcoming = upcoming;
            forTable = new ArrayList<Integer>();
            sb = new StringBuilder(1000);
        }

        /**
         * <p>Drafts a table of anonymous interviewee data in HTML format.</p>
         * @return a String representation of the plea email in HTML format.
         */
        public String doInBackground(){
            for(int row = 0; row < upcoming.length; row++){
                //edited to include upcoming[row][12].equals("0") in v1.2
                if(upcoming[row][12].equals("1") || upcoming[row][12].equals("0")){   //&& (upcoming[row][0].equals("") || upcoming[row][11].equals(""))? in case hosts found but emails not yet sent
                    forTable.add(row);
                }
            }
            sb.append("<table border=\"1\" style=\"borderStyle:solid\"><tr>");
            addHeaderItem("Date of Hosting");
            addHeaderItem("Interviewee Gender");
            addHeaderItem("Alma Mater");
            addHeaderItem("Preferred Host Gender");
            addHeaderItem("Allergies");
            addHeaderItem("Interest Groups");
            sb.append("</tr>");
            for(int i=0; i < forTable.size(); i++){
                addRow(i);
            }
            sb.append("</table>");
            return sb.toString();
        }

        /**
         * <p>Drafts the plea email containing an table of anonymous interviewee data in HTML format.</p>
         * <p>The email body is read from the Google Sheet "PritzkerHosting_PATH", sheet 1, cell B5. The
         * Merge Tags [PLEA TABLE] and [SIGNATURE] are replaced by the HTML table and the names of all the
         * hosting coordinators as specified in teh Google Sheet "PritzkerHosting_PATH", sheet 1, cell B2.</p>
         */
        public void done(){
            String emailBody = sheetsHandler.getPleaEmailText();
            String pleaTable = null;
            try {
                pleaTable = get();
            } catch (Exception e) {
                e.printStackTrace();
                pleaTable = "INSERT TABLE HERE";
            }
            emailBody = emailBody.replaceAll("\\[PLEA TABLE\\]",pleaTable);
            emailBody = emailBody.replaceAll("\\[SIGNATURE\\]",sheetsHandler.getPleaEmailSignature());
            //System.out.println(emailBody);
            askClassMessage.setText(emailBody);
            askClassMessage.setCaretPosition(0);
        }

        /**
         * Helper method to add a header item to the HTML plea table.
         * @param header the column header being added
         */
        private void addHeaderItem(String header){
            sb.append("<th>");
            sb.append(header);
            sb.append("</th>");
        }

        /**
         * Adds a row to the plea table containing relevant data from a single interviewee.
         * @param index the row number of the interviewee in the given interviewee info array
         */
        private void addRow(int index){
            int row = forTable.get(index);
            sb.append("<tr><td style=\"color:red\">");
            sb.append(upcoming[row][4]);
            sb.append("</td><td>");
            sb.append(upcoming[row][2]);
            sb.append("</td><td>");
            sb.append(upcoming[row][7]);
            sb.append("</td><td>");
            sb.append(upcoming[row][8]);
            sb.append("</td><td>");
            sb.append(upcoming[row][9]);
            sb.append("</td><td>");
            sb.append(upcoming[row][10]);
            sb.append("</td></tr>");
        }
    }

    /**
     * <p>A subclass of <code>javax.swing.SwingWorker</code> to initialize GUI components
     * and connect to the Gmail server and Google Drive off the GUI thread.</p>
     */
    private class StartupWorker extends SwingWorker<Void,Void>{
        private PritzkerEHR pritzkerEHR;

        /**
         * <p>Constructor for the <code>StartupWorker</code>.</p>
         * @param phh2 the PritzkerEHR application that is being initialized
         */
        public StartupWorker(PritzkerEHR phh2){
            this.pritzkerEHR = phh2;
        }

        /**
         * Implentation of the initialization of GUI components and connection to Gmail and Google Drive.
         * @return null (n/a)
         */
        public Void doInBackground(){
            //sets email and write progress counters to 0
            resetProgressCounters();
            //sets the progress bar to not show write progress until at least one email is sent (for aesthetics)
            suppressWriteProgress = false;
            //initializes the SheetsHandler for the application
            try {
                sheetsHandler = new SheetsHandler(pritzkerEHR);
            }
            catch(Throwable t){     //in case there was an error in initializing the SheetsHandler (usually a connection error)
                t.printStackTrace();
                JOptionPane.showMessageDialog(pritzkerEHR,"Pritzker EHR has just had trouble connecting to Google Drive. Please wait two seconds and try again.","Oops!",JOptionPane.WARNING_MESSAGE);
                System.exit(1);
            }
            //load icons
            saveIcon = new ImageIcon(getClass().getResource("resources/save_small.gif"));
            stepForwardIcon = new ImageIcon(getClass().getResource("resources/step_forward_small.gif"));
            fastForwardIcon = new ImageIcon(getClass().getResource("resources/fast_forward_small.gif"));
            sendMailIcon = new ImageIcon(getClass().getResource("resources/send_mail_small.gif"));
            //initialize the layout of the application window
            setLayout(new BorderLayout());
            //initialize the options dropdown menu at the top of the window
            messageType = new JComboBox(new String[] {CONFIRM_INT_REQUEST_RECEIVED, CONFIRM_HOST_TO_INTERVIEWEES, SEND_PLEA_TO_CLASS_LISTS, VIEW_COMPLETED_REQUESTS, VIEW_IGNORE_LIST,HELP_AND_ABOUT}); //also add SEND_PLEA_TO_CLASS_LISTS
            messageType.setSelectedItem(CONFIRM_INT_REQUEST_RECEIVED);
            messageType.addActionListener(pritzkerEHR);
            //initialize the different center panels used for displaying different tables
            requestReceivedPanel = new JPanel(new BorderLayout());
            tellIntervieweesPanel = new JPanel(new BorderLayout());
            ignoredRequestsPanel = new JPanel(new BorderLayout());
            doneIntervieweePanel = new JPanel(new BorderLayout());
            //initialize the table for interviwees awaiting their receipt email
            jtR = sheetsHandler.getReceiptJTable();
            jtR.setPreferredScrollableViewportSize(new Dimension(500,200));
            jtR.setFillsViewportHeight(true);
            jtR.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
                @Override
                public void valueChanged(ListSelectionEvent e) {
                    sheetsHandler.updateInfoArea(jtR,intervieweeInfoArea,CONFIRM_INT_REQUEST_RECEIVED);
                }
            });
            //clear the panel of previous contents
            requestReceivedPanel.removeAll();
            //place the table in a scroll pane
            JScrollPane jspR = new JScrollPane(jtR,JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
            requestReceivedPanel.add(jspR);
            //initialize the table of interviewees waiting to be paired
            jtP = sheetsHandler.getPairingJTable();
            jtP.setPreferredScrollableViewportSize(new Dimension(500, 200));
            jtP.setFillsViewportHeight(true);
            jtP.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
                @Override
                public void valueChanged(ListSelectionEvent e) {
                    sheetsHandler.updateInfoArea(jtP, intervieweeInfoArea, CONFIRM_HOST_TO_INTERVIEWEES);
                }
            });
            tellIntervieweesPanel.removeAll();
            JScrollPane jspP = new JScrollPane(jtP,JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
            tellIntervieweesPanel.add(jspP);
            //initialize the table of interviewee hosting requests that are being ignored
            jtI = sheetsHandler.getIgnoredJTable();
            jtI.setPreferredScrollableViewportSize(new Dimension(500,200));
            jtI.setFillsViewportHeight(true);
            jtI.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
                @Override
                public void valueChanged(ListSelectionEvent e) {
                    sheetsHandler.updateInfoArea(jtI,intervieweeInfoArea,VIEW_IGNORE_LIST);
                }
            });
            JScrollPane jspI = new JScrollPane(jtI,JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
            ignoredRequestsPanel.add(jspI);
            //initialize the table for interviwees who've been paired
            jtD = sheetsHandler.getDoneJTable();
            jtD.setPreferredScrollableViewportSize(new Dimension(500,200));
            jtD.setFillsViewportHeight(true);
            jtD.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
                @Override
                public void valueChanged(ListSelectionEvent e) {
                    sheetsHandler.updateInfoArea(jtD,intervieweeInfoArea,VIEW_COMPLETED_REQUESTS);
                }
            });
            //clear the panel of previous contents
            doneIntervieweePanel.removeAll();
            //place the table in a scroll pane
            JScrollPane jspD = new JScrollPane(jtD,JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
            doneIntervieweePanel.add(jspD);
            //update the plea table and format the interviewee tables
            updatePleaAndTables();
            //initialize the hosting plea panel
            askClassPanel = new JPanel();
            askClassMessage = new JTextPane();
            askClassMessage.setContentType("text/html");
            JScrollPane jspA = new JScrollPane(askClassMessage,JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
            jspA.setPreferredSize(new Dimension(682,229));
            askClassPanel.add(jspA);
            askClassButton = new JButton("Send Hosting Plea to Class Lists");
            askClassButton.setIcon(sendMailIcon);
            askClassButton.addActionListener(pritzkerEHR);
            askClassRecipientsPanel = new JPanel();
            askClassRecipientsPanel.add(new JLabel("Subject:"));
            askClassSubjectField = new JTextField();
            askClassSubjectField.setText("Pritzker Hosting_ Call for Hosts");
            askClassRecipientsPanel.add(askClassSubjectField);
            askClassRecipientsPanel.add(new JLabel("To:"));
            askClassRecipientsField = new JTextField();
            askClassRecipientsField.setText(initializePleaRecipients());
            askClassRecipientsPanel.add(askClassRecipientsField);
            //initialize center panel which will contain the more specific center panels
            centerPanel = new JPanel();
            centerPanel.add(requestReceivedPanel);
            //initialize the text area on the right side displaying interviewee details
            intervieweeInfoArea = new JTextArea("Click on an applicant to view details");
            intervieweeInfoArea.setLineWrap(true);
            intervieweeInfoArea.setWrapStyleWord(true);
            intervieweeInfoArea.setEditable(false);
            intInfoScroll = new JScrollPane(intervieweeInfoArea);
            intInfoScroll.setPreferredSize(new Dimension(170,238));
            intInfoScroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
            centerPanel.add(intInfoScroll);
            //initialize the bottom panel that will contain buttons and the signature field
            bottomPanel = new JPanel(new BorderLayout());
            //initialize the signature panel that will be placed in the bottom panel
            signatureField = new JTextField(20);
            sigPanel = new JPanel();
            sigPanel.add(new JLabel("Sender's Name:"));
            sigPanel.add(signatureField);
            bottomPanel.add(sigPanel,BorderLayout.NORTH);
            //initialize the button panel that will be placed in the bottom panel
            buttonPanel = new JPanel(new GridLayout(1,3));
            saveToDocButton = new JButton("Save Edits to Doc");
            saveToDocButton.setIcon(saveIcon); //or icon could be UIManager.getIcon("FileView.floppyDriveIcon")
            saveToDocButton.addActionListener(pritzkerEHR);
            buttonPanel.add(saveToDocButton);
            launchStepwiseButton = new JButton("Launch with Individual Review");
            launchStepwiseButton.setIcon(stepForwardIcon);
            buttonPanel.add(launchStepwiseButton,BorderLayout.WEST);
            launchStepwiseButton.addActionListener(pritzkerEHR);
            launchAllButton = new JButton("Launch All");
            launchAllButton.setIcon(fastForwardIcon);
            buttonPanel.add(launchAllButton, BorderLayout.EAST);
            launchAllButton.addActionListener(pritzkerEHR);
            bottomPanel.add(buttonPanel,BorderLayout.SOUTH);
            //initialize the progress bar that will be displayed in the bottom panel during operations
            progressBar = new JProgressBar(JProgressBar.HORIZONTAL,0,100);
            progressBar.setStringPainted(true);
            //add all top level panels to the application window
            add(messageType,BorderLayout.NORTH);
            add(centerPanel,BorderLayout.CENTER);
            add(bottomPanel,BorderLayout.SOUTH);
            //SwingWorkers need a return value by default
            return null;
        }

        /**
         * <p>Sets the progress to 100, thereby firing a <code>PropertyChangedEvent</code> indicating that the initialization is complete.</p>
         */
        public void done(){
            setProgress(100);
        }
    }
}