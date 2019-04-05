/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.databackup;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.Serializable;
import java.util.LinkedList;
import java.util.TimerTask;
import java.util.Timer;
import xlstoxml.service.Modules;

/**
 *
 * @author DCruz
 */
public class AutoSaver {
    private boolean autosaveXLS = true;
    private boolean autosaveParsingRules = true;
    private boolean autosaveRenameRules = true;
    
    private Timer timer = new Timer();
    private TimerTask task = null;
    
    private static AutoSaver instance = new AutoSaver();
    
    private AutoSaver(){}

    public static AutoSaver getInstance(){
        return instance;
    }

    public void setAutosaveXLS(boolean autosaveXLS) {
        this.autosaveXLS = autosaveXLS;
    }

    public void setAutosaveParsingRules(boolean autosaveParsingRules) {
        this.autosaveParsingRules = autosaveParsingRules;
    }

    public void setAutosaveRenameRules(boolean autosaveRenameRules) {
        this.autosaveRenameRules = autosaveRenameRules;
    }
    
    //First item in the list is the output folder, the rest is the list of xls/xlsx files
    public void serialize(String outputFolder, LinkedList<String> xlsList, Boolean renameChildren,
                          String originalName, String newName, LinkedList<LinkedList<String>> mapping,
                          LinkedList<LinkedList<String>> rulesMapping) throws Exception{
        UploadNConvert unc = new UploadNConvert();
        TagRename tr = new TagRename();
        ParsingRules dr = new ParsingRules();

        if(autosaveXLS){
            unc.outputFolder = outputFolder;
            unc.xlsList.addAll(xlsList);
        }
        
        if(autosaveRenameRules){
            tr.oldName = originalName;
            tr.newName = newName;
            tr.renameMapping = mapping;
            tr.renameChildren = renameChildren;
        }
        
        if(autosaveParsingRules){
            dr.parsingMapping = rulesMapping;
        }
        
        serializeAny("uploadNConvert", unc);
        
        serializeAny("tagRename", tr);
        
        serializeAny("parsingRules", dr);
    }
    
    public UploadNConvert deserializeXLS() throws Exception{
        Object ser = deserializeAny("uploadNConvert");
        return ser == null ? new UploadNConvert() : (UploadNConvert)ser;
    }

    ParsingRules deserializeDefineRules() throws Exception {
        Object ser = deserializeAny("parsingRules");
        return ser == null ? new ParsingRules() : (ParsingRules)ser;
    }
    
    public TagRename deserializeTagRename() throws Exception{
        Object ser = deserializeAny("tagRename");
        return ser == null ? new TagRename() : (TagRename)ser;
    }
    
    private void serializeAny(String fileName, Serializable data){
        try{
            String filepath = PropertiesManager.getPropFolder()+"\\"+fileName+".ser";
            File serFile = new File(filepath);
            
            if(serFile.exists() && serFile.canWrite()){
                FileOutputStream fileOut = null;
                fileOut = new FileOutputStream(filepath);
                ObjectOutputStream out = new ObjectOutputStream(fileOut);
                out.writeObject(data);
                out.close();
                fileOut.close();
            }
        }
        catch(Exception ex){//InvalidClassException
            Modules.getConsoleService().log(data.toString());
            Modules.getConsoleService().error(ex.getMessage());
            Modules.getConsoleService().warning("File could not be saved: "+fileName+". This may be a incompatibility problem.");
        }
    }
    
    Object deserializeAny(String fileName) throws Exception {
        Object ser = null;
        String filepath = PropertiesManager.getPropFolder()+"\\"+fileName+".ser";

        try{
            File serFile = new File(filepath);
            
            if(serFile.exists() && serFile.canRead()){
                FileInputStream fileIn = new FileInputStream(filepath);
                ObjectInputStream in = new ObjectInputStream(fileIn);
                ser = in.readObject();
                in.close();
                fileIn.close();
            }
        }
        catch(Exception ex){//InvalidClassException
            Modules.getConsoleService().error(ex.getMessage());
            for(StackTraceElement ste : ex.getStackTrace()){
                Modules.getConsoleService().error(ste.toString());
            }
            Modules.getConsoleService().warning("File could not be loaded: "+filepath+". This may be a incompatibility problem.");
        }
        
        return ser;
    }
}
