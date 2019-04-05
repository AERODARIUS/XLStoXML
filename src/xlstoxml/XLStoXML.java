/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml;

import xlstoxml.helper.CustomTableCellRenderer;
import xlstoxml.helper.SpinnerEditor;
import xlstoxml.helper.FileDrop;
import xlstoxml.helper.TableCellListener;
import xlstoxml.helper.ConvertThread;
import xlstoxml.databackup.PropertiesManager;
import xlstoxml.databackup.AutoSaver;
import java.awt.Color;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.GridBagLayout;
import java.awt.LayoutManager;
import java.awt.event.ActionEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;
import java.util.Timer;
import java.util.TimerTask;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.CellEditor;
import javax.swing.DefaultCellEditor;
import javax.swing.DefaultListModel;
import javax.swing.ImageIcon;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JEditorPane;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.SpinnerModel;
import javax.swing.SpinnerNumberModel;
import javax.swing.SwingConstants;
import javax.swing.UIDefaults;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;
import javax.swing.text.html.HTMLEditorKit;
import xlstoxml.service.IConsoleService;
import xlstoxml.service.Modules;
import xlstoxml.service.IDefineRulesService;
import xlstoxml.service.ITagRenameService;
import xlstoxml.service.IUploadNConvertService;
import xlstoxml.helper.*;

/**
 *
 * @author ASUS
 */
public class XLStoXML extends javax.swing.JFrame {
    IDefineRulesService drTab = Modules.getDefineRulesService();
    
    ITagRenameService renameRules = Modules.getTagRenameService();
    
    IUploadNConvertService converter = Modules.getUploadNConvertService();
    
    IConsoleService console = Modules.getConsoleService();
    
    
    private int counter = 0;

    DefaultListModel<String> dlm = new DefaultListModel<String>();

    DefaultListModel<LinkedList<String>> dlm3 = new DefaultListModel<LinkedList<String>>();
    
    SpinnerModel modelSpinner1 = new SpinnerNumberModel(1, 1, 100.0000, 1);
    SpinnerModel modelSpinner2 = new SpinnerNumberModel(1, 1, 100.0000, 1);
    DefaultTableModel tableModel = new DefaultTableModel();
    
    
    //Create a file chooser
    private JFileChooser XLSChooser = null;
    private JFileChooser FolderChooser = null;
    private JFileChooser PropertyChooser = null;
    private JFileChooser LogChooser = null;
    
    //private String message = "";
    
    private int rowPos = 0;
    
    private int colPos = 0;
    
    private String sheetName = "";
    
    private String workbookName = "";
    
    private HashMap<Integer,HashMap<Integer,Boolean>> disabledCells = new HashMap<Integer,HashMap<Integer,Boolean>>();
    
    private FileDrop fd = null;
    
    private HTMLEditorKit kit = new HTMLEditorKit();
    
    private HashSet<String> filesDirsSet = new HashSet<String>();
        
    private HashMap<String, JComponent> settingComponents = new HashMap<String, JComponent>();
    
    HashMap<Integer, Long> frequencyMapping = new HashMap<Integer, Long>();
    
    private Timer timer = new Timer();
    private TimerTask task = null;
    
    private JLabel gears = new JLabel("",
                                      new javax.swing.ImageIcon(getClass().getResource("/img/gears.gif")),
                                      JLabel.CENTER);
    
            
    /**
     * Creates new form XLStoXML
     */
    public XLStoXML() {
        try {
            frequencyMapping.put(1, new Long(1000*30));
            frequencyMapping.put(2, new Long(1000*60));
            frequencyMapping.put(3, new Long(1000*60*2));
            frequencyMapping.put(4, new Long(1000*60*3));
            frequencyMapping.put(5, new Long(1000*60*5));
            frequencyMapping.put(6, new Long(1000*60*8));
            frequencyMapping.put(7, new Long(1000*60*13));
            frequencyMapping.put(8, new Long(1000*60*21));
            frequencyMapping.put(9, new Long(1000*60*34));
            frequencyMapping.put(10, new Long(1000*60*55));
            
            initComponents();
            
            console.addLogPane(logPane);
            
            jPanel10.add(gears);

            gears.setVisible(false);
        
            //this.setExtendedState(this.getExtendedState() | JFrame.MAXIMIZED_BOTH);
            setIconImage(new ImageIcon(getClass().getResource("/img/icon_ligth.png")).getImage());
            tableModel.addColumn("Apply Rule To");
            tableModel.addColumn("Order");
            tableModel.addColumn("Name");
            tableModel.addColumn("Enconding");
            tableModel.addColumn("Reorder");
            tableModel.addColumn("Is Main Wrapper");
            tableModel.addColumn("Has Header");
            tableModel.addColumn("Cell Type");
            tableModel.addColumn("Match By");
            
            jTable1.setDefaultRenderer(Object.class, new CustomTableCellRenderer(disabledCells));
            jTable1.getTableHeader().setReorderingAllowed(false);
            
            TableColumn tCol = jTable1.getColumnModel().getColumn(0);
            String [] options1 = {"XLS File","XLS Tab","Column/Row"};
            JComboBox comboBox = new JComboBox<String>(options1);
            tCol.setCellEditor(new DefaultCellEditor(comboBox));
            
            tCol = jTable1.getColumnModel().getColumn(1);
            tCol.setCellEditor(new SpinnerEditor());
            
            tCol = jTable1.getColumnModel().getColumn(3);
            //https://docs.oracle.com/javase/7/docs/technotes/guides/intl/encoding.doc.html
            //http://w3techs.com/technologies/overview/character_encoding/all
            
            Set<String> chSet = java.nio.charset.Charset.availableCharsets().keySet();
            String [] options2 = new String [chSet.size()];
            int k = 0;
            
            for(String chs : chSet){
                options2[k++] = chs;
            }
                                /*{""+StandardCharsets.UTF_8,
                                  ""+StandardCharsets.ISO_8859_1,
                                  ""+StandardCharsets.US_ASCII,
                                  ""+StandardCharsets.UTF_16,
                                  ""+StandardCharsets.UTF_16BE,
                                  ""+StandardCharsets.UTF_16LE};*/
            
            comboBox = new JComboBox<String>(options2);
            tCol.setCellEditor(new DefaultCellEditor(comboBox));
            
            tCol = jTable1.getColumnModel().getColumn(4);
            tCol.setCellEditor(new SpinnerEditor());
            
            tCol = jTable1.getColumnModel().getColumn(5);
            String [] options3 = {"no", "yes"};
            comboBox = new JComboBox<String>(options3);
            tCol.setCellEditor(new DefaultCellEditor(comboBox));
            
            tCol = jTable1.getColumnModel().getColumn(6);
            comboBox = new JComboBox<String>(options3);
            tCol.setCellEditor(new DefaultCellEditor(comboBox));
            
            tCol = jTable1.getColumnModel().getColumn(7);
            String [] options4 = {"Text", "CDATA", "Property", "Reference", "SubTag"};
            comboBox = new JComboBox<String>(options4);
            tCol.setCellEditor(new DefaultCellEditor(comboBox));
            
            tCol = jTable1.getColumnModel().getColumn(8);
            String [] options5 = {"Columns", "Rows"};
            comboBox = new JComboBox<String>(options5);
            tCol.setCellEditor(new DefaultCellEditor(comboBox));
            
            Action action = new AbstractAction(){
                @Override
                public void actionPerformed(ActionEvent e){
                    try{
                        TableCellListener tcl = (TableCellListener)e.getSource();
                        
                        int ruleNumber = tcl.getRow();
                        String colName = jTable1.getColumnName(tcl.getColumn());
                        String value = ""+tcl.getNewValue();
                        
                        if("Apply Rule To".equals(colName)){
                            setColumnVisibility(value, ruleNumber);
                        }
                        
                        String [] allValues = new String[jTable1.getColumnCount()];
                        
                        for(int k = 0; k < allValues.length; k++){
                            allValues[k] = ""+jTable1.getModel().getValueAt(tcl.getRow(), k);
                        }
                        
                        drTab.setRuleValue(ruleNumber, colName, value, allValues);
                    }
                    catch (Exception ex) {
                        Logger.getLogger(XLStoXML.class.getName()).log(Level.SEVERE, null, ex);
                    }
                }
            };
            
            TableCellListener tcl = new TableCellListener(jTable1, action);
            jTable1.addPropertyChangeListener(tcl);
            
            fd = new FileDrop(System.out, jList1, /*dragBorder,*/ new FileDrop.Listener(){
                @Override
                public void filesDropped( java.io.File[] files ){
                    try {
                        LinkedList<File> selectedFiles = new LinkedList<File>();
                        boolean tryToReaddFiles = false;
                        
                        for(File file : files){
                            try {
                                String path = file.getCanonicalPath();
                                String ext = path.indexOf("\\") == -1? "":
                                        path.substring(path.lastIndexOf("\\")+1);
                                
                                ext = path.indexOf(".") == -1? "":
                                        path.substring(path.lastIndexOf(".")+1);
                                
                                if("xls".equals(ext) || "xlsx".equals(ext)){
                                    if(!filesDirsSet.contains(path)){
                                        dlm.addElement(path);
                                        filesDirsSet.add(path);
                                    }
                                    else{
                                        console.warning("File with path "+path+" was already added.");
                                        tryToReaddFiles = true;
                                    }
                                    
                                    selectedFiles.addLast(file);
                                }
                            }catch( java.io.IOException e ){
                                console.error("Failed to load file "+file.getName());
                            }
                        }
                        
                        if(tryToReaddFiles){
                            console.warning("Some files were not added. If you are trying to update files with modified content, press refresh button instead.", uploadNConvertMessages);
                        }
                        else{
                            console.succes("All files added successfully!",uploadNConvertMessages);
                        }
                        
                        File[] XLSfiles = new File[selectedFiles.size()];
                        int k = 0;
                        
                        for(File f : selectedFiles){
                            XLSfiles[k++] = f;
                        }
                        
                        converter.addFiles(XLSfiles);
                    } catch (Exception ex) {
                        console.error("Failed to drop files.");
                    }
                }
            });
            
            //Init File choosers
            XLSChooser = new JFileChooser();
            XLSChooser.setAcceptAllFileFilterUsed(false);
            XLSChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            XLSChooser.setMultiSelectionEnabled(true);
            XLSChooser.setFileFilter(new FileNameExtensionFilter("xls and xlsx", new String[] { "xls", "xlsx" }));
            
            String currentDirectory;
            File file = new File("");
            currentDirectory = file.getAbsolutePath();
            jLabel21.setText(currentDirectory);
            
            FolderChooser = new JFileChooser(currentDirectory);
            FolderChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
            
            PropertyChooser = new JFileChooser();
            PropertyChooser.setAcceptAllFileFilterUsed(false);
            PropertyChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            PropertyChooser.setMultiSelectionEnabled(false);
            PropertyChooser.setFileFilter(new FileNameExtensionFilter("properties", new String[] { "properties" }));
            
            LogChooser = new JFileChooser();
            LogChooser.setAcceptAllFileFilterUsed(false);
            LogChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
            LogChooser.setMultiSelectionEnabled(false);
            LogChooser.setFileFilter(new FileNameExtensionFilter("log", new String[] { "log" }));
            
            UIDefaults labelTable = new UIDefaults();
            labelTable.put(0, new JLabel("No"));
            labelTable.put(1, new JLabel("Yes"));
            autosaveRenameRules.setLabelTable(labelTable);
            autosaveExcelList.setLabelTable(labelTable);
            autosaveParsingRules.setLabelTable(labelTable);
            
            labelTable = new UIDefaults();
            labelTable.put(1, new JLabel("30 sec"));
            labelTable.put(2, new JLabel("1 min"));
            labelTable.put(3, new JLabel("2 min"));
            labelTable.put(4, new JLabel("3 min"));
            labelTable.put(5, new JLabel("5 min"));
            labelTable.put(6, new JLabel("8 min"));
            labelTable.put(7, new JLabel("13 min"));
            labelTable.put(8, new JLabel("21 min"));
            labelTable.put(9, new JLabel("34 min"));
            labelTable.put(10, new JLabel("55 min"));
            labelTable.put(11, new JLabel("On Close"));
            labelTable.put(12, new JLabel("Never"));
            autosaveFrequency.setLabelTable(labelTable);
            
            settingComponents.put("autosaveExcelList",autosaveExcelList);
            settingComponents.put("autosaveParsingRules",autosaveParsingRules);
            settingComponents.put("autosaveRenameRules",autosaveRenameRules);
            settingComponents.put("autosaveFrequency",autosaveFrequency);
            
            settingComponents.put("autosaveFileLogs",autosaveFileLogs);
            
            settingComponents.put("smartConversion",smartConversion);
            
            PropertiesManager.load(settingComponents);
            
            jLabel21.setText(PropertiesManager.outputFolder);
            LinkedList<File> xlsToLoad = new LinkedList<File>();
            
            for(String dir : PropertiesManager.xlsList){
                dlm.addElement(dir);
                filesDirsSet.add(dir);
                File xls = new File(dir);
                xlsToLoad.add(xls);
            }
            
            File[] xlsArray = new File[xlsToLoad.size()];
            
            converter.addFiles(xlsToLoad.toArray(xlsArray));
            
            jTextField1.setText(PropertiesManager.originalName);
            
            jTextField5.setText(PropertiesManager.newName);
            
            jCheckBox1.setSelected(PropertiesManager.renameChildren);
            
            for(LinkedList<String> entry : PropertiesManager.renameMapping){
                if(entry.size() > 2){
                    renameRules.addRule(entry.get(0), entry.get(1), Boolean.parseBoolean(entry.get(2)));
                }
                
                dlm3.addElement(entry);
            }
            
            for(LinkedList<String> row : PropertiesManager.parsingMapping){
                //tableModel
                String ruleType = row.get(0);
                Integer order = Integer.parseInt(row.get(1));
                String name = row.get(2);
                String enconding = row.get(3);
                Integer reorder = Integer.parseInt(row.get(4));
                String isMainWrapper = row.get(5);
                String hasHeader = row.get(6);
                String cellType = row.get(7);
                String matchBy = row.get(8);
                    
                String result = drTab.addFormatingRule(matchBy,
                                                        enconding,
                                                        reorder,
                                                        "yes".equals(hasHeader),
                                                        "yes".equals(isMainWrapper),
                                                        cellType,
                                                        ruleType,
                                                        order,
                                                        name);
            
                if(!result.contains("Error")){
                    Object [] rowToInsert = {ruleType, order, name, enconding, reorder,
                                             isMainWrapper, hasHeader, cellType,matchBy};

                    tableModel.insertRow(jTable1.getRowCount(), rowToInsert);

                    setColumnVisibility(ruleType, tableModel.getRowCount()-1);

                    console.succes(result);
                }
                else{
                    console.error(result);
                }
            }
            
            if(autosaveFrequency.getValue() < 11){
                task = new TimerTask(){
                    @Override
                    public void run(){
                        try {
                            PropertiesManager.store(settingComponents);
                            console.log("Autosave done.");
                        } catch (Exception ex) {
                            console.error("Autosave could not be performed in this moment.");
                        }
                    }
                };
                
                timer.schedule(task, 10, frequencyMapping.get(autosaveFrequency.getValue()));
            }
        } catch (Exception ex) {
            //Logger.getLogger(XLStoXML.class.getName()).log(Level.SEVERE, null, ex);
            console.error("Unknown error when loading.");
            console.error(ex.getMessage());//TODO: Remove
        }
    }
    
    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jCheckBox2 = new javax.swing.JCheckBox();
        jTabbedPane1 = new javax.swing.JTabbedPane();
        jScrollPane9 = new javax.swing.JScrollPane();
        jPanel5 = new javax.swing.JPanel();
        jLabel6 = new javax.swing.JLabel();
        jLabel1 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jScrollPane3 = new javax.swing.JScrollPane();
        jPanel1 = new javax.swing.JPanel();
        jPanel10 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        jList1 = new javax.swing.JList();
        jButton3 = new javax.swing.JButton();
        jButton5 = new javax.swing.JButton();
        jLabel3 = new javax.swing.JLabel();
        jButton4 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        uploadNConvertMessages = new javax.swing.JLabel();
        jLabel21 = new javax.swing.JLabel();
        jButton8 = new javax.swing.JButton();
        jScrollPane4 = new javax.swing.JScrollPane();
        jPanel2 = new javax.swing.JPanel();
        jLabel11 = new javax.swing.JLabel();
        jLabel20 = new javax.swing.JLabel();
        jButton7 = new javax.swing.JButton();
        jButton6 = new javax.swing.JButton();
        jLabel16 = new javax.swing.JLabel();
        jScrollPane6 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable(){
            public boolean isCellEditable(int row, int column){
                if(disabledCells.get(row).get(column)){
                    return false;
                }
                return true;
            }
        };
        jButton9 = new javax.swing.JButton();
        jButton10 = new javax.swing.JButton();
        jScrollPane7 = new javax.swing.JScrollPane();
        jPanel4 = new javax.swing.JPanel();
        jScrollPane8 = new javax.swing.JScrollPane();
        jList3 = new javax.swing.JList();
        jList3.setModel(dlm3);
        jButton12 = new javax.swing.JButton();
        jButton11 = new javax.swing.JButton();
        jButton13 = new javax.swing.JButton();
        jTextField5 = new javax.swing.JTextField();
        jLabel14 = new javax.swing.JLabel();
        jLabel17 = new javax.swing.JLabel();
        jCheckBox1 = new javax.swing.JCheckBox();
        jTextField1 = new javax.swing.JTextField();
        jLabel15 = new javax.swing.JLabel();
        jLabel19 = new javax.swing.JLabel();
        jButton14 = new javax.swing.JButton();
        jButton15 = new javax.swing.JButton();
        jButton16 = new javax.swing.JButton();
        jPanel3 = new javax.swing.JPanel();
        jScrollPane2 = new javax.swing.JScrollPane();
        logPane = new javax.swing.JTextPane();
        jButton18 = new javax.swing.JButton();
        jPanel7 = new javax.swing.JPanel();
        jLabel9 = new javax.swing.JLabel();
        jPanel9 = new javax.swing.JPanel();
        jLabel12 = new javax.swing.JLabel();
        jLabel13 = new javax.swing.JLabel();
        jScrollPane11 = new javax.swing.JScrollPane();
        jTextArea3 = new javax.swing.JTextArea();
        jLabel7 = new javax.swing.JLabel();
        jScrollPane5 = new javax.swing.JScrollPane();
        jPanel8 = new javax.swing.JPanel();
        jLabel2 = new javax.swing.JLabel();
        autosaveRenameRules = new javax.swing.JSlider();
        jLabel10 = new javax.swing.JLabel();
        autosaveFileLogs = new javax.swing.JTextField();
        jButton17 = new javax.swing.JButton();
        jLabel24 = new javax.swing.JLabel();
        jLabel25 = new javax.swing.JLabel();
        autosaveExcelList = new javax.swing.JSlider();
        autosaveParsingRules = new javax.swing.JSlider();
        jLabel27 = new javax.swing.JLabel();
        jLabel26 = new javax.swing.JLabel();
        autosaveFrequency = new javax.swing.JSlider();
        smartConversion = new javax.swing.JCheckBox();

        jCheckBox2.setText("jCheckBox2");

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("XLS to XML 2.0");
        setAlwaysOnTop(true);
        setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        setIconImages(null);
        addWindowListener(new java.awt.event.WindowAdapter() {
            public void windowClosing(java.awt.event.WindowEvent evt) {
                formWindowClosing(evt);
            }
        });

        jTabbedPane1.setBackground(new java.awt.Color(204, 204, 255));
        jTabbedPane1.setTabPlacement(javax.swing.JTabbedPane.LEFT);

        jScrollPane9.setEnabled(false);

        jPanel5.setBackground(new java.awt.Color(249, 104, 84));
        jPanel5.setPreferredSize(new java.awt.Dimension(183, 182));

        jLabel6.setFont(new java.awt.Font("Arial", 1, 36)); // NOI18N
        jLabel6.setForeground(new java.awt.Color(5, 45, 73));
        jLabel6.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel6.setText("Support this project!");

        jLabel1.setFont(new java.awt.Font("Arial", 0, 14)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(0, 102, 204));
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("http://xlstoxml.sourceforge.net/donate");
        jLabel1.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        jLabel1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel4MouseClicked(evt);
            }
        });

        jLabel4.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel4.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/donate-paypal.png"))); // NOI18N
        jLabel4.setCursor(new java.awt.Cursor(java.awt.Cursor.WAIT_CURSOR));
        jLabel4.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel4MouseClicked(evt);
            }
        });

        javax.swing.GroupLayout jPanel5Layout = new javax.swing.GroupLayout(jPanel5);
        jPanel5.setLayout(jPanel5Layout);
        jPanel5Layout.setHorizontalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jLabel4, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel6, javax.swing.GroupLayout.DEFAULT_SIZE, 664, Short.MAX_VALUE)
                .addContainerGap())
            .addComponent(jLabel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        jPanel5Layout.setVerticalGroup(
            jPanel5Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel5Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel6, javax.swing.GroupLayout.PREFERRED_SIZE, 63, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(115, 115, 115)
                .addComponent(jLabel4)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel1)
                .addContainerGap(223, Short.MAX_VALUE))
        );

        jScrollPane9.setViewportView(jPanel5);

        jTabbedPane1.addTab("Support                 ", new javax.swing.ImageIcon(getClass().getResource("/img/iconfinder_donate_59557.png")), jScrollPane9); // NOI18N

        jPanel1.setBackground(new java.awt.Color(204, 204, 255));
        jPanel1.setPreferredSize(new java.awt.Dimension(685, 484));

        jScrollPane1.setAutoscrolls(true);

        jList1.setBackground(new java.awt.Color(240, 240, 240));
        jList1.setModel(dlm);
        jList1.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        jScrollPane1.setViewportView(jList1);

        javax.swing.GroupLayout jPanel10Layout = new javax.swing.GroupLayout(jPanel10);
        jPanel10.setLayout(jPanel10Layout);
        jPanel10Layout.setHorizontalGroup(
            jPanel10Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane1, javax.swing.GroupLayout.Alignment.TRAILING)
        );
        jPanel10Layout.setVerticalGroup(
            jPanel10Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jScrollPane1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, 316, Short.MAX_VALUE)
        );

        jButton3.setFont(new java.awt.Font("Arial", 1, 24)); // NOI18N
        jButton3.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/upload.png"))); // NOI18N
        jButton3.setToolTipText("Add xls/xlsx files");
        jButton3.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton3ActionPerformed(evt);
            }
        });

        jButton5.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/folder.png"))); // NOI18N
        jButton5.setToolTipText("Change output folder");
        jButton5.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton5ActionPerformed(evt);
            }
        });

        jLabel3.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        jLabel3.setForeground(new java.awt.Color(51, 51, 51));
        jLabel3.setText("Output Folder:");

        jButton4.setFont(new java.awt.Font("Tahoma", 0, 24)); // NOI18N
        jButton4.setForeground(new java.awt.Color(204, 0, 0));
        jButton4.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/remove.png"))); // NOI18N
        jButton4.setToolTipText("Remove selected files");
        jButton4.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton4ActionPerformed(evt);
            }
        });

        jButton2.setFont(new java.awt.Font("Tahoma", 0, 24)); // NOI18N
        jButton2.setForeground(new java.awt.Color(0, 102, 51));
        jButton2.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/convert.png"))); // NOI18N
        jButton2.setToolTipText("Convert all or the selected ones");
        jButton2.setPreferredSize(new java.awt.Dimension(64, 64));
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        uploadNConvertMessages.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        uploadNConvertMessages.setForeground(new java.awt.Color(51, 51, 51));
        uploadNConvertMessages.setText("Successful conversions:  0");

        jLabel21.setOpaque(true);

        jButton8.setFont(new java.awt.Font("Tahoma", 0, 24)); // NOI18N
        jButton8.setForeground(new java.awt.Color(0, 102, 51));
        jButton8.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/refresh.png"))); // NOI18N
        jButton8.setToolTipText("Refresh list and file contents");
        jButton8.setPreferredSize(new java.awt.Dimension(64, 64));
        jButton8.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton8ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jPanel10, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(uploadNConvertMessages, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jLabel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(jLabel21, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(javax.swing.GroupLayout.Alignment.LEADING, jPanel1Layout.createSequentialGroup()
                        .addComponent(jButton3)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jButton4)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 94, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jButton8, javax.swing.GroupLayout.PREFERRED_SIZE, 94, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jButton5)
                        .addGap(0, 162, Short.MAX_VALUE)))
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jPanel10, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jButton3)
                            .addComponent(jButton5)
                            .addComponent(jButton2, javax.swing.GroupLayout.PREFERRED_SIZE, 73, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton8, javax.swing.GroupLayout.PREFERRED_SIZE, 73, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabel3)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabel21, javax.swing.GroupLayout.PREFERRED_SIZE, 21, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(uploadNConvertMessages))
                    .addComponent(jButton4))
                .addContainerGap())
        );

        jScrollPane3.setViewportView(jPanel1);

        jTabbedPane1.addTab("Upload & Convert", new javax.swing.ImageIcon(getClass().getResource("/img/flask.png")), jScrollPane3); // NOI18N

        jPanel2.setBackground(new java.awt.Color(204, 204, 255));
        jPanel2.setPreferredSize(new java.awt.Dimension(284, 287));

        jLabel11.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N

        jLabel20.setFont(new java.awt.Font("Tahoma", 3, 11)); // NOI18N
        jLabel20.setForeground(new java.awt.Color(51, 102, 255));
        jLabel20.setText("In this section the parsing rules are defined.");

        jButton7.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/add.png"))); // NOI18N
        jButton7.setToolTipText("Add new parsing rule");
        jButton7.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton7ActionPerformed(evt);
            }
        });

        jButton6.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/remove2.png"))); // NOI18N
        jButton6.setToolTipText("Remove selected rules");
        jButton6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton6ActionPerformed(evt);
            }
        });

        jLabel16.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        jLabel16.setText("Parsing Rules :");

        jTable1.setModel(tableModel);
        jScrollPane6.setViewportView(jTable1);

        jButton9.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/arrow_up.png"))); // NOI18N
        jButton9.setToolTipText("Move rules up");
        jButton9.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton9ActionPerformed(evt);
            }
        });

        jButton10.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/arrow_down.png"))); // NOI18N
        jButton10.setToolTipText("Move rules down");
        jButton10.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton10ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel20, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel2Layout.createSequentialGroup()
                        .addComponent(jScrollPane6, javax.swing.GroupLayout.DEFAULT_SIZE, 660, Short.MAX_VALUE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabel11))
                    .addComponent(jLabel16, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addComponent(jButton7, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jButton6, javax.swing.GroupLayout.PREFERRED_SIZE, 44, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jButton10, javax.swing.GroupLayout.PREFERRED_SIZE, 44, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(jButton9, javax.swing.GroupLayout.PREFERRED_SIZE, 44, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap())
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addGap(25, 25, 25)
                .addComponent(jLabel16)
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addGap(249, 249, 249)
                        .addComponent(jLabel11)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 127, Short.MAX_VALUE))
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jScrollPane6, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)))
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jButton7)
                    .addComponent(jButton6)
                    .addComponent(jButton10)
                    .addComponent(jButton9))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel20, javax.swing.GroupLayout.PREFERRED_SIZE, 26, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap())
        );

        jScrollPane4.setViewportView(jPanel2);

        jTabbedPane1.addTab("Define Rules        ", new javax.swing.ImageIcon(getClass().getResource("/img/filter.png")), jScrollPane4); // NOI18N

        jPanel4.setBackground(new java.awt.Color(204, 204, 255));
        jPanel4.setPreferredSize(new java.awt.Dimension(269, 296));

        jList3.setBackground(new java.awt.Color(240, 240, 240));
        jList3.setSelectionMode(javax.swing.ListSelectionModel.SINGLE_SELECTION);
        jScrollPane8.setViewportView(jList3);

        jButton12.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/add.png"))); // NOI18N
        jButton12.setToolTipText("Add new rename rule");
        jButton12.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton12ActionPerformed(evt);
            }
        });

        jButton11.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/info.png"))); // NOI18N
        jButton11.setToolTipText("Load rename rule info");
        jButton11.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton11ActionPerformed(evt);
            }
        });

        jButton13.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/remove2.png"))); // NOI18N
        jButton13.setToolTipText("Remove selected rename rule");
        jButton13.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton13ActionPerformed(evt);
            }
        });

        jTextField5.setToolTipText("Specify the replacement");

        jLabel14.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        jLabel14.setText("Mapping :");

        jLabel17.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        jLabel17.setText("New Name:");

        jCheckBox1.setBackground(new java.awt.Color(204, 204, 255));
        jCheckBox1.setText("Rename Children");
        jCheckBox1.setToolTipText("If checked the immediate child nodes are renamed");

        jTextField1.setToolTipText("Specify the original tag name");

        jLabel15.setFont(new java.awt.Font("Tahoma", 1, 11)); // NOI18N
        jLabel15.setHorizontalAlignment(javax.swing.SwingConstants.RIGHT);
        jLabel15.setText("Original Name:");

        jLabel19.setFont(new java.awt.Font("Tahoma", 2, 11)); // NOI18N
        jLabel19.setForeground(new java.awt.Color(51, 102, 255));
        jLabel19.setText("In this section you can rename tags.");

        jButton14.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/arrow_down.png"))); // NOI18N
        jButton14.setToolTipText("Move rename rule down");
        jButton14.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton14ActionPerformed(evt);
            }
        });

        jButton15.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/arrow_up.png"))); // NOI18N
        jButton15.setToolTipText("Move rename rule up");
        jButton15.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton15ActionPerformed(evt);
            }
        });

        jButton16.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/edit.png"))); // NOI18N
        jButton16.setToolTipText("Update rename rule info");
        jButton16.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton16ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel4Layout = new javax.swing.GroupLayout(jPanel4);
        jPanel4.setLayout(jPanel4Layout);
        jPanel4Layout.setHorizontalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jScrollPane8)
                            .addComponent(jLabel14, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addGroup(jPanel4Layout.createSequentialGroup()
                                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addComponent(jLabel19, javax.swing.GroupLayout.PREFERRED_SIZE, 452, javax.swing.GroupLayout.PREFERRED_SIZE)
                                    .addGroup(jPanel4Layout.createSequentialGroup()
                                        .addGap(132, 132, 132)
                                        .addComponent(jButton12, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(18, 18, 18)
                                        .addComponent(jButton11, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(18, 18, 18)
                                        .addComponent(jButton16, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(18, 18, 18)
                                        .addComponent(jButton13, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(18, 18, 18)
                                        .addComponent(jButton14, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(18, 18, 18)
                                        .addComponent(jButton15, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE)))
                                .addGap(0, 154, Short.MAX_VALUE)))
                        .addContainerGap())
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addGap(24, 24, 24)
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel15, javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(jLabel17, javax.swing.GroupLayout.Alignment.TRAILING))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jTextField5, javax.swing.GroupLayout.DEFAULT_SIZE, 221, Short.MAX_VALUE)
                            .addComponent(jTextField1))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jCheckBox1, javax.swing.GroupLayout.PREFERRED_SIZE, 117, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
        );
        jPanel4Layout.setVerticalGroup(
            jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel4Layout.createSequentialGroup()
                .addGap(18, 18, 18)
                .addComponent(jLabel19)
                .addGap(12, 12, 12)
                .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel15)
                            .addComponent(jTextField1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jCheckBox1))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel17)
                            .addComponent(jTextField5, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addGroup(jPanel4Layout.createSequentialGroup()
                        .addGap(101, 101, 101)
                        .addGroup(jPanel4Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jButton12, javax.swing.GroupLayout.PREFERRED_SIZE, 43, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton11, javax.swing.GroupLayout.PREFERRED_SIZE, 43, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton13, javax.swing.GroupLayout.PREFERRED_SIZE, 43, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton16, javax.swing.GroupLayout.PREFERRED_SIZE, 43, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton14, javax.swing.GroupLayout.PREFERRED_SIZE, 43, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton15, javax.swing.GroupLayout.PREFERRED_SIZE, 43, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabel14)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jScrollPane8, javax.swing.GroupLayout.DEFAULT_SIZE, 281, Short.MAX_VALUE)))
                .addContainerGap())
        );

        jScrollPane7.setViewportView(jPanel4);

        jTabbedPane1.addTab("Tag Rename       ", new javax.swing.ImageIcon(getClass().getResource("/img/tag.png")), jScrollPane7); // NOI18N

        jPanel3.setCursor(new java.awt.Cursor(java.awt.Cursor.TEXT_CURSOR));

        jScrollPane2.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(153, 153, 153)));

        logPane.setEditable(false);
        logPane.setBackground(new java.awt.Color(0, 0, 0));
        logPane.setBorder(javax.swing.BorderFactory.createEtchedBorder());
        logPane.setContentType("text/html"); // NOI18N
        logPane.setEditorKit(kit);
        logPane.setFont(new java.awt.Font("Arial", 1, 11)); // NOI18N
        logPane.setToolTipText("");
        logPane.setCursor(new java.awt.Cursor(java.awt.Cursor.TEXT_CURSOR));
        jScrollPane2.setViewportView(logPane);

        jButton18.setText("Clear Console");
        jButton18.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        jButton18.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                removeLogs(evt);
            }
        });

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addGap(0, 587, Short.MAX_VALUE)
                .addComponent(jButton18))
            .addComponent(jScrollPane2)
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 472, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jButton18, javax.swing.GroupLayout.PREFERRED_SIZE, 23, javax.swing.GroupLayout.PREFERRED_SIZE))
        );

        jTabbedPane1.addTab("Console              ", new javax.swing.ImageIcon(getClass().getResource("/img/console.png")), jPanel3); // NOI18N

        jLabel9.setFont(new java.awt.Font("Tahoma", 0, 18)); // NOI18N
        jLabel9.setText("Ways to get support:");

        jPanel9.setBackground(new java.awt.Color(48, 48, 48));

        jLabel12.setFont(new java.awt.Font("Arial", 0, 48)); // NOI18N
        jLabel12.setForeground(new java.awt.Color(255, 173, 0));
        jLabel12.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel12.setText("XLS to XML 2.0");
        jLabel12.setHorizontalTextPosition(javax.swing.SwingConstants.RIGHT);

        jLabel13.setFont(new java.awt.Font("Arial", 0, 24)); // NOI18N
        jLabel13.setForeground(new java.awt.Color(204, 204, 204));
        jLabel13.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel13.setText("Translate files between XLS / XLSX and XML format . . .");
        jLabel13.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));

        javax.swing.GroupLayout jPanel9Layout = new javax.swing.GroupLayout(jPanel9);
        jPanel9.setLayout(jPanel9Layout);
        jPanel9Layout.setHorizontalGroup(
            jPanel9Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jLabel12, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addComponent(jLabel13, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        jPanel9Layout.setVerticalGroup(
            jPanel9Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel9Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel12)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel13, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(25, Short.MAX_VALUE))
        );

        jScrollPane11.setBorder(null);

        jTextArea3.setEditable(false);
        jTextArea3.setBackground(new java.awt.Color(240, 240, 240));
        jTextArea3.setColumns(20);
        jTextArea3.setRows(5);
        jTextArea3.setText("Copyright 2018 XLStoXML\n\nLicensed under the Apache License, Version 2.0 (the \"License\");\nyou may not use this file except in compliance with the License.\nYou may obtain a copy of the License at\n\n    http://www.apache.org/licenses/LICENSE-2.0\n\nUnless required by applicable law or agreed to in writing, software\ndistributed under the License is distributed on an \"AS IS\" BASIS,\nWITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.\nSee the License for the specific language governing permissions and\nlimitations under the License.");
        jTextArea3.setWrapStyleWord(true);
        jScrollPane11.setViewportView(jTextArea3);

        jLabel7.setFont(new java.awt.Font("Tahoma", 0, 18)); // NOI18N
        jLabel7.setForeground(new java.awt.Color(0, 0, 153));
        jLabel7.setText("http://xlstoxml.sourceforge.net/#Support");
        jLabel7.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jLabel7.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                goToSupport(evt);
            }
        });

        javax.swing.GroupLayout jPanel7Layout = new javax.swing.GroupLayout(jPanel7);
        jPanel7.setLayout(jPanel7Layout);
        jPanel7Layout.setHorizontalGroup(
            jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel9, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addGroup(jPanel7Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel9)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel7)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
            .addGroup(jPanel7Layout.createSequentialGroup()
                .addGap(10, 10, 10)
                .addComponent(jScrollPane11))
        );
        jPanel7Layout.setVerticalGroup(
            jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel7Layout.createSequentialGroup()
                .addComponent(jPanel9, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane11, javax.swing.GroupLayout.PREFERRED_SIZE, 244, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(39, 39, 39)
                .addGroup(jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel7)
                    .addComponent(jLabel9))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        jTabbedPane1.addTab("Help                    ", new javax.swing.ImageIcon(getClass().getResource("/img/support.png")), jPanel7); // NOI18N

        jScrollPane5.setBorder(null);

        jLabel2.setFont(new java.awt.Font("Arial", 1, 11)); // NOI18N
        jLabel2.setText("Autosave excel file list:");

        autosaveRenameRules.setMaximum(1);
        autosaveRenameRules.setMinorTickSpacing(1);
        autosaveRenameRules.setPaintLabels(true);
        autosaveRenameRules.setPaintTicks(true);
        autosaveRenameRules.setSnapToTicks(true);
        autosaveRenameRules.setAutoscrolls(true);
        autosaveRenameRules.setFocusable(false);
        autosaveRenameRules.addChangeListener(new javax.swing.event.ChangeListener() {
            public void stateChanged(javax.swing.event.ChangeEvent evt) {
                autosaveRenameRulesStateChanged(evt);
            }
        });

        jLabel10.setFont(new java.awt.Font("Arial", 1, 11)); // NOI18N
        jLabel10.setText("Log file:");

        autosaveFileLogs.setEditable(false);
        autosaveFileLogs.setBackground(new java.awt.Color(255, 255, 255));
        autosaveFileLogs.setCursor(new java.awt.Cursor(java.awt.Cursor.TEXT_CURSOR));
        autosaveFileLogs.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                autosaveFileLogsMouseClicked(evt);
            }
        });

        jButton17.setIcon(new javax.swing.ImageIcon(getClass().getResource("/img/property-directory.png"))); // NOI18N
        jButton17.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton17ActionPerformed(evt);
            }
        });

        jLabel24.setFont(new java.awt.Font("Arial", 1, 11)); // NOI18N
        jLabel24.setText("Autosave parsing rules:");

        jLabel25.setFont(new java.awt.Font("Arial", 1, 11)); // NOI18N
        jLabel25.setText("Autosave rename rules:");

        autosaveExcelList.setMaximum(1);
        autosaveExcelList.setMinorTickSpacing(1);
        autosaveExcelList.setPaintLabels(true);
        autosaveExcelList.setPaintTicks(true);
        autosaveExcelList.setSnapToTicks(true);
        autosaveExcelList.setToolTipText("");
        autosaveExcelList.setAutoscrolls(true);
        autosaveExcelList.setFocusable(false);
        autosaveExcelList.addChangeListener(new javax.swing.event.ChangeListener() {
            public void stateChanged(javax.swing.event.ChangeEvent evt) {
                autosaveExcelListStateChanged(evt);
            }
        });

        autosaveParsingRules.setMaximum(1);
        autosaveParsingRules.setMinorTickSpacing(1);
        autosaveParsingRules.setPaintLabels(true);
        autosaveParsingRules.setPaintTicks(true);
        autosaveParsingRules.setSnapToTicks(true);
        autosaveParsingRules.setAutoscrolls(true);
        autosaveParsingRules.setFocusable(false);
        autosaveParsingRules.addChangeListener(new javax.swing.event.ChangeListener() {
            public void stateChanged(javax.swing.event.ChangeEvent evt) {
                autosaveParsingRulesStateChanged(evt);
            }
        });

        jLabel27.setText(" ");

        jLabel26.setFont(new java.awt.Font("Arial", 1, 11)); // NOI18N
        jLabel26.setText("Autosave frequency:");

        autosaveFrequency.setMaximum(12);
        autosaveFrequency.setMinimum(1);
        autosaveFrequency.setMinorTickSpacing(1);
        autosaveFrequency.setPaintLabels(true);
        autosaveFrequency.setPaintTicks(true);
        autosaveFrequency.setSnapToTicks(true);
        autosaveFrequency.setFocusable(false);
        autosaveFrequency.addChangeListener(new javax.swing.event.ChangeListener() {
            public void stateChanged(javax.swing.event.ChangeEvent evt) {
                autosaveFrequencyStateChanged(evt);
            }
        });

        smartConversion.setFont(new java.awt.Font("Arial", 1, 11)); // NOI18N
        smartConversion.setText("Smart conversion (Fast)");

        javax.swing.GroupLayout jPanel8Layout = new javax.swing.GroupLayout(jPanel8);
        jPanel8.setLayout(jPanel8Layout);
        jPanel8Layout.setHorizontalGroup(
            jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel8Layout.createSequentialGroup()
                .addGroup(jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel8Layout.createSequentialGroup()
                        .addGap(10, 10, 10)
                        .addGroup(jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel8Layout.createSequentialGroup()
                                .addComponent(jLabel10)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(autosaveFileLogs)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jButton17, javax.swing.GroupLayout.PREFERRED_SIZE, 48, javax.swing.GroupLayout.PREFERRED_SIZE))
                            .addGroup(jPanel8Layout.createSequentialGroup()
                                .addComponent(jLabel26)
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(autosaveFrequency, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                            .addGroup(jPanel8Layout.createSequentialGroup()
                                .addGroup(jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                    .addGroup(jPanel8Layout.createSequentialGroup()
                                        .addComponent(jLabel24)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                        .addComponent(autosaveParsingRules, javax.swing.GroupLayout.PREFERRED_SIZE, 74, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(18, 18, 18)
                                        .addComponent(smartConversion))
                                    .addGroup(jPanel8Layout.createSequentialGroup()
                                        .addComponent(jLabel2)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                        .addComponent(autosaveExcelList, javax.swing.GroupLayout.PREFERRED_SIZE, 74, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addGap(18, 18, 18)
                                        .addComponent(jLabel25)
                                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                                        .addComponent(autosaveRenameRules, javax.swing.GroupLayout.PREFERRED_SIZE, 74, javax.swing.GroupLayout.PREFERRED_SIZE)))
                                .addGap(0, 223, Short.MAX_VALUE))))
                    .addGroup(jPanel8Layout.createSequentialGroup()
                        .addContainerGap()
                        .addComponent(jLabel27, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                .addContainerGap())
        );
        jPanel8Layout.setVerticalGroup(
            jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel8Layout.createSequentialGroup()
                .addGap(18, 18, 18)
                .addGroup(jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(autosaveExcelList, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel25, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(autosaveRenameRules, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(20, 20, 20)
                .addGroup(jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabel24, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(autosaveParsingRules, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(smartConversion))
                .addGroup(jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel8Layout.createSequentialGroup()
                        .addGap(30, 30, 30)
                        .addComponent(jLabel26, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel8Layout.createSequentialGroup()
                        .addGap(18, 18, 18)
                        .addComponent(autosaveFrequency, javax.swing.GroupLayout.PREFERRED_SIZE, 65, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGroup(jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel8Layout.createSequentialGroup()
                        .addGap(34, 34, 34)
                        .addComponent(jButton17, javax.swing.GroupLayout.PREFERRED_SIZE, 38, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel8Layout.createSequentialGroup()
                        .addGap(44, 44, 44)
                        .addGroup(jPanel8Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(autosaveFileLogs, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel10))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jLabel27)
                .addContainerGap())
        );

        jScrollPane5.setViewportView(jPanel8);

        jTabbedPane1.addTab("Settings              ", new javax.swing.ImageIcon(getClass().getResource("/img/settings.png")), jScrollPane5); // NOI18N

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jTabbedPane1)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jTabbedPane1)
        );

        setSize(new java.awt.Dimension(848, 545));
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents
    
    private LayoutManager previousLayout = null;
    
    private LayoutManager showOverlay(){
        jScrollPane1.setVisible(false);
        jPanel10.setBackground(new Color(0xf0f0f0));
        LayoutManager layoutManager = jPanel10.getLayout();
        jPanel10.setLayout(new GridBagLayout());
        gears.setVisible(true);
        jPanel10.repaint();
        
        return layoutManager;
    }
    
        
    private void openlogFile(JTextField field){
        try{
            LogChooser.setCurrentDirectory(new File(field.getText()));
            
            int returnVal = LogChooser.showOpenDialog(this);

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File props = LogChooser.getSelectedFile();
                autosaveFileLogs.setText(props.getAbsolutePath());
            }
        }
        catch(Exception e){
            console.error("Failed to open log file.", jLabel27);
        }
    }
        
    private void formWindowClosing(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowClosing
        try {
            if(task != null){
                task.cancel();
            }
            
            PropertiesManager.outputFolder = jLabel21.getText();
            PropertiesManager.xlsList = new LinkedList<String>();
            
            for(int k = 0; k < jList1.getModel().getSize(); k++){
                String dir = (String)jList1.getModel().getElementAt(k);
                PropertiesManager.xlsList.add(dir);
            }
            
            PropertiesManager.originalName = jTextField1.getText();
            PropertiesManager.newName = jTextField5.getText();
            PropertiesManager.renameChildren = jCheckBox1.isSelected();
            PropertiesManager.renameMapping = new LinkedList<LinkedList<String>>();
            PropertiesManager.parsingMapping = new LinkedList<LinkedList<String>>();
            
            for(int k = 0; k < dlm3.size(); k++){
                LinkedList<String> entry = (LinkedList<String>) dlm3.getElementAt(k);
                PropertiesManager.renameMapping.add(entry);
            }
                 
            for(int k = 0; k < tableModel.getRowCount(); k++){
                LinkedList<String> cells = new LinkedList<String>();
                
                for(int n = 0; n < tableModel.getColumnCount(); n++){
                    String val = ""+tableModel.getValueAt(k, n);
                    cells.add(val);
                }
                
                PropertiesManager.parsingMapping.add(cells);
            }
            
            PropertiesManager.store(settingComponents);
        } catch (Exception ex) {
            console.error("Error when loading properties.");
        }
    }//GEN-LAST:event_formWindowClosing

    private void autosaveFrequencyStateChanged(javax.swing.event.ChangeEvent evt) {//GEN-FIRST:event_autosaveFrequencyStateChanged
        if(autosaveFrequency.getValue() < 11){
            if(task != null){
                task.cancel();
            }

            task = new TimerTask(){
                @Override
                public void run(){
                    try {
                        PropertiesManager.store(settingComponents);
                        console.log("Autosave done.");
                    } catch (Exception ex) {
                        console.error("Autosave didn't complete.");
                    }
                }
            };

            timer.schedule(task, 10, frequencyMapping.get(autosaveFrequency.getValue()));
        }
    }//GEN-LAST:event_autosaveFrequencyStateChanged

    private void autosaveParsingRulesStateChanged(javax.swing.event.ChangeEvent evt) {//GEN-FIRST:event_autosaveParsingRulesStateChanged
        int val = autosaveParsingRules.getValue();
        AutoSaver as = AutoSaver.getInstance();
        as.setAutosaveParsingRules(val == 1);
    }//GEN-LAST:event_autosaveParsingRulesStateChanged

    private void autosaveExcelListStateChanged(javax.swing.event.ChangeEvent evt) {//GEN-FIRST:event_autosaveExcelListStateChanged
        int val = autosaveExcelList.getValue();
        AutoSaver as = AutoSaver.getInstance();
        as.setAutosaveXLS(val == 1);
    }//GEN-LAST:event_autosaveExcelListStateChanged

    private void jButton17ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton17ActionPerformed
        openlogFile(autosaveFileLogs);
    }//GEN-LAST:event_jButton17ActionPerformed

    private void autosaveFileLogsMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_autosaveFileLogsMouseClicked
        openlogFile(autosaveFileLogs);
    }//GEN-LAST:event_autosaveFileLogsMouseClicked

    private void autosaveRenameRulesStateChanged(javax.swing.event.ChangeEvent evt) {//GEN-FIRST:event_autosaveRenameRulesStateChanged
        int val = autosaveRenameRules.getValue();
        AutoSaver as = AutoSaver.getInstance();
        as.setAutosaveRenameRules(val == 1);
    }//GEN-LAST:event_autosaveRenameRulesStateChanged

    private void goToSupport(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_goToSupport
        openWebpage("http://xlstoxml.sourceforge.net/#Support");
    }//GEN-LAST:event_goToSupport

    private void removeLogs(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_removeLogs
        console.clear();
    }//GEN-LAST:event_removeLogs

    private void jButton16ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton16ActionPerformed
        try{
            Integer ind = jList3.getSelectedIndex();

            if(ind != -1){
                String originalName = jTextField1.getText().replaceAll("\\s+","");
                String newName = jTextField5.getText().replaceAll("\\s+","");
                Boolean renameChildren = jCheckBox1.isSelected();

                LinkedList<String> entry = new LinkedList<String>();
                entry.add(originalName);
                entry.add(newName);
                entry.add(renameChildren.toString());

                if(originalName != null && !originalName.isEmpty() &&
                    newName != null && !newName.isEmpty()){
                    renameRules.updateRule(ind, originalName, newName, renameChildren);
                    dlm3.set(ind, entry);
                }

                console.succes("Rename rule updated successfully!", jLabel19);
            }
        }
        catch(Exception e){
            console.error("Rule override did not complete. Check if you have not already added the rule.", jLabel19);
        }
    }//GEN-LAST:event_jButton16ActionPerformed

    private void jButton15ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton15ActionPerformed
        try{
            Integer ind = jList3.getSelectedIndex();

            if(ind != -1 && ind > 0){
                renameRules.moveUp(ind);
                LinkedList<String> entr1 = (LinkedList<String>)dlm3.get(ind);
                LinkedList<String> entry2 = (LinkedList<String>)dlm3.get(ind-1);
                dlm3.set(ind, entry2);
                dlm3.set(ind-1, entr1);
                jList3.setSelectedIndex(ind-1);
            }

            console.succes("Rename rule moved successfully!", jLabel19);
        }
        catch(Exception e){
            console.error("Rule rename didn't complete.", jLabel19);
        }
    }//GEN-LAST:event_jButton15ActionPerformed

    private void jButton14ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton14ActionPerformed
        try{
            Integer ind = jList3.getSelectedIndex();

            if(ind != -1 && ind < dlm3.size()-1){
                renameRules.moveDown(ind);
                LinkedList<String> entr1 = (LinkedList<String>)dlm3.get(ind);
                LinkedList<String> entry2 = (LinkedList<String>)dlm3.get(ind+1);
                dlm3.set(ind, entry2);
                dlm3.set(ind+1, entr1);
                jList3.setSelectedIndex(ind+1);
            }

            console.succes("Rename rule moved successfully!", jLabel19);
        }
        catch(Exception e){
            console.error("Failed to move rule.", jLabel19);
        }
    }//GEN-LAST:event_jButton14ActionPerformed

    private void jButton13ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton13ActionPerformed
        try{
            Integer ind = jList3.getSelectedIndex();

            if(ind != -1){
                renameRules.removeRule(ind);
                dlm3.remove(ind);

                console.succes("Rename rule removed successfully!", jLabel19);
            }
        }
        catch(Exception e){
            console.error("Failed to remove rule.", jLabel19);
        }
    }//GEN-LAST:event_jButton13ActionPerformed

    private void jButton11ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton11ActionPerformed
        try{
            Integer ind = jList3.getSelectedIndex();

            if(ind != -1){
                LinkedList<String> entry = (LinkedList<String>)dlm3.getElementAt(ind);
                jTextField1.setText(entry.get(0));
                jTextField5.setText(entry.get(1));
                jCheckBox1.setSelected("true".equals(entry.get(2)));
            }
        }
        catch(Exception e){
            console.error("Failed to load rename rule.", jLabel19);
        }
    }//GEN-LAST:event_jButton11ActionPerformed

    private void jButton12ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton12ActionPerformed
        try{
            String originalName = jTextField1.getText().replaceAll("\\s+","");
            String newName = jTextField5.getText().replaceAll("\\s+","");
            Boolean renameChildren = jCheckBox1.isSelected();

            if(!originalName.isEmpty() && !newName.isEmpty()){
                LinkedList<String> entry = new LinkedList<String>();
                entry.add(originalName);
                entry.add(newName);
                entry.add(renameChildren.toString());

                renameRules.addRule(originalName, newName, renameChildren);
                dlm3.addElement(entry);

                console.succes("Rename rule added successfully!", jLabel19);
            }
        }
        catch(Exception e){
            console.error("Failed to add rename rule.  Check if you have not already added the rule.", jLabel19);
        }
    }//GEN-LAST:event_jButton12ActionPerformed

    private void jButton10ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton10ActionPerformed
        try{
            moveRows(1);//Foward
            console.succes("Rules moved succesfully", jLabel20);
        }
        catch(Exception ex){
            console.error("Failed to move rules.", jLabel20);
        }
        catch(Error er){
            console.error("Failed to move rules.", jLabel20);
        }
    }//GEN-LAST:event_jButton10ActionPerformed

    private void jButton9ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton9ActionPerformed
        try{
            moveRows(-1);//Backward
            console.succes("Rules moved succesfully", jLabel20);
        }
        catch(Exception ex){
            console.error("Failed to move rules.", jLabel20);
        }
        catch(Error er){
            console.error("Failed to move rules.", jLabel20);
        }
    }//GEN-LAST:event_jButton9ActionPerformed

    private void jButton6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton6ActionPerformed
        int [] selectedRows = jTable1.getSelectedRows();

        if(selectedRows.length > 0){
            try{
                drTab.removeFormatingRule(selectedRows);

                stopCellEditing();

                for(int k = selectedRows.length-1;k > -1; k--){
                    tableModel.removeRow(selectedRows[k]);
                    removeColumnsVisibility(selectedRows[k]);
                }

                console.succes("Rules removed succesfully", jLabel20);
            }
            catch(Exception ex){
                console.error("Failed to remove rules.", jLabel20);
            }
            catch(Error er){
                console.error("Failed to remove rules.", jLabel20);
            }
        }
        else{
            console.warning("No rows selected", jLabel20);
        }
    }//GEN-LAST:event_jButton6ActionPerformed

    private void jButton7ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton7ActionPerformed
        try {
            String ruleType = "XLS File";
            String matchBy = "Rows";
            String enconding = "UTF-8";
            Integer order = 0;
            Integer reorder = 0;
            Boolean isMainWrapper = true;
            Boolean hasHeader = true;
            String cellType = "Text";
            String name = "";

            String result = drTab.addFormatingRule(matchBy,
                enconding,
                reorder,
                hasHeader,
                isMainWrapper,
                cellType,
                ruleType,
                order,
                name);

            if(!result.contains("Error")){
                Object [] row = {ruleType, order, name, enconding, reorder,
                    isMainWrapper?"yes":"no", hasHeader?"yes":"no",
                    cellType,matchBy};

                tableModel.insertRow(jTable1.getRowCount(), row);

                setColumnVisibility(ruleType, tableModel.getRowCount()-1);

                console.succes(result, jLabel20);
            }
            else{
                console.error(result, jLabel20);
            }
        } catch (Exception ex) {
            console.error(ex.toString(), jLabel20);
        }
    }//GEN-LAST:event_jButton7ActionPerformed

    private void jButton8ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton8ActionPerformed
        try{
            //Reload files content or remove deleted files
            HashSet<String> paths = converter.updateFileList();
            filesDirsSet = new HashSet<String>();
            LinkedList<String> allPaths = new LinkedList<String>();

            for(int k = 0;k<dlm.getSize();k++){
                allPaths.add((String) dlm.get(k));
            }

            dlm.removeAllElements();

            for(String p : allPaths){
                if(paths.contains(p)){
                    dlm.addElement(p);
                    filesDirsSet.add(p);
                }
                else{
                    console.log("File "+p+" was removed from the list!");
                }
            }
        }
        catch(Exception e){
            console.error("An error occurred while changing output directory.", uploadNConvertMessages);
        }
    }//GEN-LAST:event_jButton8ActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        previousLayout = showOverlay();
        ConvertThread cth = new ConvertThread(smartConversion.isSelected(), console, converter, jList1, previousLayout, jLabel21,
            uploadNConvertMessages, gears, jPanel10, jScrollPane1);
        cth.start();
        try {
            cth.waitUntilFinish();
        } catch (Exception ex) {
            jList1.setBackground(Color.RED);
            console.error("Convert didn't succeed.");
            //System.out.println(ex.getStackTrace());
        }
    }//GEN-LAST:event_jButton2ActionPerformed

    private void jButton4ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton4ActionPerformed
        LinkedList<String> removedSuccess = new LinkedList<String>();

        try{
            int[] indices = jList1.getSelectedIndices();

            for(int k = 0;k<indices.length;k++) {
                String path = (String) dlm.getElementAt(indices[k]-(k));
                filesDirsSet.remove(path);
                dlm.remove(indices[k]-(k));
                removedSuccess.add("File "+path+" removed successfully.");
            }

            converter.removeFiles(indices);

            for(String msg : removedSuccess){
                console.succes(msg);
            }

            console.succes("All files where removed successfully",uploadNConvertMessages);
        }
        catch(Exception e){
            console.error("Failed to remove files.", uploadNConvertMessages);
        }
    }//GEN-LAST:event_jButton4ActionPerformed

    private void jButton5ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton5ActionPerformed
        try{
            int returnVal = FolderChooser.showOpenDialog(this);

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                File folder = FolderChooser.getSelectedFile();
                jLabel21.setText(folder.getAbsolutePath());
                //This is where a real application would open the file.
                //log.append("Opening: " + file.getName() + "." + newline);
            }
        }
        catch(Exception e){
            console.error("An error occurred while changing output directory.", uploadNConvertMessages);
        }
    }//GEN-LAST:event_jButton5ActionPerformed

    private void jButton3ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton3ActionPerformed
        int returnVal = XLSChooser.showOpenDialog(this);
        boolean tryToReaddFiles = false;

        if (returnVal == JFileChooser.APPROVE_OPTION) {
            try {
                File[] selectedFiles = XLSChooser.getSelectedFiles();

                for(File XLSfile : selectedFiles){
                    String absolutePath = XLSfile.getAbsolutePath();
                    console.log(absolutePath);
                    if(!filesDirsSet.contains(absolutePath)){
                        dlm.addElement(absolutePath);
                        filesDirsSet.add(absolutePath);
                    }
                    else{
                        console.warning("File with path "+absolutePath+" was already added.");
                        tryToReaddFiles = true;
                    }
                }

                converter.addFiles(selectedFiles);

                if(tryToReaddFiles){
                    console.warning("Some files were not added. If you are trying to update files with modified content, press refresh button instead.", uploadNConvertMessages);
                }
                else{
                    console.succes("All files added successfully!",uploadNConvertMessages);
                }
            } catch (Exception ex) {
                console.error(ex.getMessage());
                console.error("Failed to open files.");
            }
        }
    }//GEN-LAST:event_jButton3ActionPerformed

    private void jLabel4MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel4MouseClicked
        openWebpage("http://xlstoxml.sourceforge.net/donate");
    }//GEN-LAST:event_jLabel4MouseClicked

    private Object [] getAllComboboxItems(JComboBox myComboBox){
        Object [] items = new Object[myComboBox.getItemCount()];
        
        for(int i = 0; i < myComboBox.getItemCount(); i++){
            items[i] = myComboBox.getItemAt(i);
        }
        
        return items;
    }
    
    private void removeColumnsVisibility(Integer rowNum){
        disabledCells.remove(rowNum);
        
        List<Integer> orderedKyes = new ArrayList<Integer>(disabledCells.keySet());
        orderedKyes.sort(null);
        
        for(int k : orderedKyes){
            if(k > rowNum){
                HashMap<Integer,Boolean> cols = disabledCells.get(k);
                disabledCells.remove(k);
                disabledCells.put(k-1, cols);
            }
        }
        
        //printVisibility();
    }
    
    private void swapRowVisibility(Integer rowNum1, Integer rowNum2){
        HashMap<Integer,Boolean> cols1 = disabledCells.get(rowNum1);
        HashMap<Integer,Boolean> cols2 = disabledCells.get(rowNum2);
        disabledCells.put(rowNum1, cols2);
        disabledCells.put(rowNum2, cols1);
        
        //printVisibility();
    }
    
    private void setColumnVisibility(String rowType, Integer rowNum){
        
        HashMap<Integer,Boolean> cols = disabledCells.containsKey(rowNum)?
                                        disabledCells.get(rowNum):
                                        new HashMap<Integer,Boolean>();
        switch(rowType){
            case "XLS File":
                cols.put(0, false);
                cols.put(1, false);
                cols.put(2, false);
                cols.put(3, false);
                cols.put(4, true);
                cols.put(5, false);
                cols.put(6, false);
                cols.put(7, true);
                cols.put(8, false);
                
            break;
            case "XLS Tab":
                cols.put(0, false);
                cols.put(1, false);
                cols.put(2, false);
                cols.put(3, true);
                cols.put(4, false);
                cols.put(5, false);
                cols.put(6, false);
                cols.put(7, true);
                cols.put(8, false);
            break;
            case "Column/Row":
                cols.put(0, false);
                cols.put(1, false);
                cols.put(2, false);
                cols.put(3, true);
                cols.put(4, true);
                cols.put(5, true);
                cols.put(6, true);
                cols.put(7, false);
                cols.put(8, true);
            break;
        }
        
        disabledCells.put(rowNum,cols);
        
        //printVisibility();
    }
    
    void printVisibility(){
        for(Integer k1 : disabledCells.keySet()){
                System.out.print(k1+" |");
            for(Integer k2 :disabledCells.get(k1).keySet()){
                System.out.print("  "+disabledCells.get(k1).get(k2));
            }
            System.out.println();
            System.out.println();
        }
        
        System.out.println("------------------------------------------------");
    }
    
    private void moveRows(int direction){
        
        int [] selectedRows = jTable1.getSelectedRows();
        
        if(selectedRows.length > 0 && (direction == 1 || direction == -1)){
            int [] reverseRows = new int[selectedRows.length];
            int [] reposition = new int[jTable1.getRowCount()];
            
            for(int r = 0;r < reposition.length;r++){
                reposition[r] = r;
            }
            
            jTable1.removeRowSelectionInterval(0, jTable1.getRowCount()-1);
        
            stopCellEditing();
            
            if(direction > 0){
                for(int r = 0;r < selectedRows.length;r++){
                    reverseRows[selectedRows.length-r-1] = selectedRows[r];
                }
            }
            else{
                reverseRows = selectedRows;
            }
            
            int positionStack = (direction<0)?0:jTable1.getRowCount()-1;
            
            for(int r = 0;r < reverseRows.length;r++){
                int pos = reverseRows[r];
                
                if((direction > 0 && pos < positionStack) || (direction < 0 && pos > positionStack)){
                    tableModel.moveRow(pos, pos, pos+direction);
                    jTable1.addRowSelectionInterval(pos+direction, pos+direction);
                    int swap = reposition[pos];
                    reposition[pos] = reposition[pos+direction];
                    reposition[pos+direction] = swap;
                    swapRowVisibility(pos,pos+direction);
                }
                else{
                    reposition[pos] = positionStack;//Not necesary?
                    jTable1.addRowSelectionInterval(positionStack, positionStack);
                    positionStack-=direction;
                }
            }
            
            drTab.formatReposition(reposition);
        }
    }
    
    private void stopCellEditing(){
        CellEditor cellEditor = jTable1.getCellEditor();

        if (cellEditor != null) {
            if (cellEditor.getCellEditorValue() != null) {
                cellEditor.stopCellEditing();
            } else {
                cellEditor.cancelCellEditing();
            }
        }
    }
    
    public static void openWebpage(String urlString) {
        try {
            Desktop.getDesktop().browse(new URL(urlString).toURI());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        try{
            /* Set the Nimbus look and feel */
            //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
            /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
             * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
             */
            try {
                /*for (LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                    if ("Metal".equals(info.getName())) {
                        UIManager.setLookAndFeel(info.getClassName());
                        break;
                    }
                    System.out.println(info.getName());
                }*/
                //UIManager.setLookAndFeel(UIManager.getCrossPlatformLookAndFeelClassName());
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
                java.util.logging.Logger.getLogger(XLStoXML.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
            }
            //</editor-fold>

            /* Create and display the form */
            java.awt.EventQueue.invokeLater(new Runnable() {
                @Override
                public void run() {
                    new XLStoXML().setVisible(true);
                }
            });
        }
        catch(Exception e){
            Modules.getConsoleService().error("Severe error when starting program.");
        }
    }
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JSlider autosaveExcelList;
    private javax.swing.JTextField autosaveFileLogs;
    private javax.swing.JSlider autosaveFrequency;
    private javax.swing.JSlider autosaveParsingRules;
    private javax.swing.JSlider autosaveRenameRules;
    private javax.swing.JButton jButton10;
    private javax.swing.JButton jButton11;
    private javax.swing.JButton jButton12;
    private javax.swing.JButton jButton13;
    private javax.swing.JButton jButton14;
    private javax.swing.JButton jButton15;
    private javax.swing.JButton jButton16;
    private javax.swing.JButton jButton17;
    private javax.swing.JButton jButton18;
    private javax.swing.JButton jButton2;
    private javax.swing.JButton jButton3;
    private javax.swing.JButton jButton4;
    private javax.swing.JButton jButton5;
    private javax.swing.JButton jButton6;
    private javax.swing.JButton jButton7;
    private javax.swing.JButton jButton8;
    private javax.swing.JButton jButton9;
    private javax.swing.JCheckBox jCheckBox1;
    private javax.swing.JCheckBox jCheckBox2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel14;
    private javax.swing.JLabel jLabel15;
    private javax.swing.JLabel jLabel16;
    private javax.swing.JLabel jLabel17;
    private javax.swing.JLabel jLabel19;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel20;
    private javax.swing.JLabel jLabel21;
    private javax.swing.JLabel jLabel24;
    private javax.swing.JLabel jLabel25;
    private javax.swing.JLabel jLabel26;
    private javax.swing.JLabel jLabel27;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JList jList1;
    private javax.swing.JList jList3;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel10;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JPanel jPanel5;
    private javax.swing.JPanel jPanel7;
    private javax.swing.JPanel jPanel8;
    private javax.swing.JPanel jPanel9;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane11;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JScrollPane jScrollPane4;
    private javax.swing.JScrollPane jScrollPane5;
    private javax.swing.JScrollPane jScrollPane6;
    private javax.swing.JScrollPane jScrollPane7;
    private javax.swing.JScrollPane jScrollPane8;
    private javax.swing.JScrollPane jScrollPane9;
    private javax.swing.JTabbedPane jTabbedPane1;
    private javax.swing.JTable jTable1;
    private javax.swing.JTextArea jTextArea3;
    private javax.swing.JTextField jTextField1;
    private javax.swing.JTextField jTextField5;
    private javax.swing.JTextPane logPane;
    private javax.swing.JCheckBox smartConversion;
    private javax.swing.JLabel uploadNConvertMessages;
    // End of variables declaration//GEN-END:variables
}
