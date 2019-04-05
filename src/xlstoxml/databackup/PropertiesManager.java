/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.databackup;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.Properties;
import javax.swing.JCheckBox;
import javax.swing.JComponent;
import javax.swing.JSlider;
import javax.swing.JTextField;
import xlstoxml.service.Modules;

/**
 *
 * @author DCruz
 */
public class PropertiesManager {
    
    public static String newLine = System.getProperty("line.separator");
    public static String userHome = System.getProperty("user.home");
    
    //Other UI Properties
    public static String outputFolder = "";
    public static LinkedList<String> xlsList = new LinkedList<String>();
    
    public static String originalName = "";
    public static String newName = "";
    public static Boolean renameChildren = false;
    public static LinkedList<LinkedList<String>> renameMapping = new LinkedList<LinkedList<String>>();
    public static LinkedList<LinkedList<String>> parsingMapping = new LinkedList<LinkedList<String>>();
    
    public static void createFileIfNotExists(String fileDir) throws Exception{
        File props = new File(fileDir);

        if(!props.exists()) {
            props.createNewFile();
            FileOutputStream oFile = new FileOutputStream(props, false);
            oFile.close();
        }
    }
    
    public static String getPropFolder() throws Exception{
        String propertiesPath = userHome+"\\XLStoXML";
                
        File archive = new File(propertiesPath);

        if(!archive.exists()){
            archive.mkdirs();
        }
        
        createFileIfNotExists(propertiesPath+"\\config.properties");
        
        createFileIfNotExists(propertiesPath+"\\autosaveFileLogs.log");
        
        return archive.getAbsolutePath();
    }
    
    private static Properties createDefaults() throws Exception{
        Properties defaults = new Properties();
        
        defaults.setProperty("autosaveExcelList", "1");
        defaults.setProperty("autosaveParsingRules", "1");
        defaults.setProperty("autosaveRenameRules", "1");
        defaults.setProperty("autosaveFrequency", "11");
        defaults.setProperty("smartConversion", "1");
        
        defaults.setProperty("autosaveFileLogs", getPropFolder()+"\\autosaveFileLogs.log");
        
        return defaults;
    }
    
    public static synchronized void load(HashMap<String, JComponent> settingComponents) throws Exception {
        Properties defaults = createDefaults();
        Properties prop = new Properties(defaults);
        
        InputStream input = new FileInputStream(getPropFolder()+"\\config.properties");
        prop.load(input);
        input.close();
        
        updateBlankProsWithDefaults(prop, defaults);
        
        createFileIfNotExists(prop.getProperty("autosaveFileLogs"));
        
        LinkedList<String> jcomponents = new LinkedList<String>();
        jcomponents.addLast("autosaveExcelList");
        jcomponents.addLast("autosaveParsingRules");
        jcomponents.addLast("autosaveRenameRules");
        jcomponents.addLast("autosaveFrequency");
        
        for(String sliderName : jcomponents){
            String p = prop.getProperty(sliderName);
            JSlider jsl = (JSlider)settingComponents.get(sliderName);
            jsl.setValue(Integer.valueOf(p));
        }
        
        jcomponents = new LinkedList<String>();
        jcomponents.addLast("autosaveFileLogs");
        
        for(String textFieldName : jcomponents){
            String dir = prop.getProperty(textFieldName);
            JTextField jtf = (JTextField)settingComponents.get(textFieldName);
            jtf.setText(dir);
        }
        
        jcomponents = new LinkedList<String>();
        jcomponents.addLast("smartConversion");
        
        for(String checkboxName : jcomponents){
            String val = prop.getProperty(checkboxName);
            JCheckBox jcb = (JCheckBox)settingComponents.get(checkboxName);
            jcb.setSelected("1".equals(val));
        }
        
        //Load UI data
        AutoSaver as = AutoSaver.getInstance();
        as.setAutosaveXLS("1".equals(prop.getProperty("autosaveExcelList")));
        as.setAutosaveParsingRules("1".equals(prop.getProperty("autosaveParsingRules")));
        as.setAutosaveRenameRules("1".equals(prop.getProperty("autosaveRenameRules")));
        
        UploadNConvert unc = as.deserializeXLS();
        outputFolder = unc.outputFolder == null || "".equals(unc.outputFolder) ? 
                       userHome:unc.outputFolder;
        xlsList = unc.xlsList;
        
        ParsingRules dr = as.deserializeDefineRules();
        parsingMapping = dr.parsingMapping;
        
        TagRename tr = as.deserializeTagRename();
        originalName = tr.oldName;
        newName = tr.newName;
        renameMapping = tr.renameMapping;
        renameChildren = tr.renameChildren;
    }
    
    public static synchronized void store(HashMap<String, JComponent> settingComponents) throws Exception{
        //Save general properties
        Properties prop = new Properties(createDefaults());
        
        LinkedList<String> jsliders = new LinkedList<String>();
        jsliders.addLast("autosaveExcelList");
        jsliders.addLast("autosaveParsingRules");
        jsliders.addLast("autosaveRenameRules");
        jsliders.addLast("autosaveFrequency");
        
        for(String sliderName : jsliders){
            JSlider jsl = (JSlider)settingComponents.get(sliderName);
            prop.setProperty(sliderName, ""+jsl.getValue());
        }
        
        jsliders = new LinkedList<String>();
        jsliders.addLast("autosaveFileLogs");
        
        for(String sliderName : jsliders){
            JTextField jtf = (JTextField)settingComponents.get(sliderName);
            prop.setProperty(sliderName, ""+jtf.getText());
        }
        
        jsliders = new LinkedList<String>();
        jsliders.addLast("smartConversion");
        
        for(String checkboxName : jsliders){
            JCheckBox jcb = (JCheckBox)settingComponents.get(checkboxName);
            prop.setProperty(checkboxName, jcb.isSelected()?"1":"0");
        }
                
        OutputStream output = new FileOutputStream(getPropFolder()+"\\config.properties");
        prop.store(output, "#General XLStoXML Properties");
        
        //Save logs to log file
        JTextField autosaveFileLogs = (JTextField)settingComponents.get("autosaveFileLogs");
        String fileName = autosaveFileLogs.getText();
        String content = Modules.getConsoleService().getMessagesToStore();
        Files.write(Paths.get(fileName), content.getBytes(), StandardOpenOption.APPEND);
        
        //Store UI data
        AutoSaver as = AutoSaver.getInstance();
        as.serialize(outputFolder, xlsList, renameChildren, originalName, newName, renameMapping, parsingMapping);
    }

    private static void updateBlankProsWithDefaults(Properties props, Properties defaults) {
        Enumeration e = props.propertyNames();
        
        while(e.hasMoreElements()){
            String key = (String) e.nextElement();

            if(defaults.contains(key) && props.getProperty(key) == null || "".equals(props.getProperty(key))){
                props.setProperty(key, defaults.getProperty(key));
            }
        }
    }
}
