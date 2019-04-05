/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic.data;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStreamWriter;
import java.nio.charset.Charset;
import java.util.HashMap;
import java.util.LinkedList;
import xlstoxml.databackup.PropertiesManager;
import xlstoxml.service.IConsoleService;
import xlstoxml.service.IDefineRulesService;
import xlstoxml.service.Modules;
import xlstoxml.service.logic.rules.RenameCollection;
import xlstoxml.service.logic.rules.XLSRule;
import org.apache.poi.ss.usermodel.*;

/**
 *
 * @author ASUS
 */
public class XLS{
    public File XLSfile;
    public Workbook sheets;
    public LinkedList<Tab> tabList = new LinkedList<Tab>();
    private HashMap<String,Tab> tabMap = new HashMap<String,Tab>();
    private XLSRule xlsRule;
    
    public XLS(File file) throws Exception{
        //Modules.getConsoleService().log("New XLS "+file.getName());
        XLSfile = file;
        
        sheets = WorkbookFactory.create(XLSfile);
        
        //Modules.getConsoleService().log("SHEETs SIZE: "+sheets.getNumberOfSheets());
        for (int k = 0; k < sheets.getNumberOfSheets(); k++){
            Sheet s = sheets.getSheetAt(k);
            Modules.getConsoleService().log("k=: "+k);
            Tab t = new Tab(this,s,k);
            tabList.add(t);
            tabMap.put(s.getSheetName(),t);
        }
    }

    void reFresh() throws Exception {
        sheets = WorkbookFactory.create(XLSfile);
        clear();
    }
    
    private class DoubleString{
        public String name = null;
        public String xmlContent = null;

        private DoubleString(String name, String tabResult) {
            this.name = name;
            xmlContent = tabResult;
        }
    }

    public LinkedList<BufferedWriter> convert(String outputDirectory, Integer order, boolean isSmartConvert) throws Exception{
        Modules.getConsoleService().log("Convert XLS "+XLSfile.getName());
        LinkedList<BufferedWriter> result = new LinkedList<BufferedWriter>();
        IDefineRulesService drs = Modules.getDefineRulesService();
        HashMap<Integer,LinkedList<DoubleString>> tabResults = new HashMap<Integer,LinkedList<DoubleString>>();
        RenameCollection rc = RenameCollection.getInstance();
        String fileName = XLSfile.getName().replaceFirst("[.][^.]+$", "");
        
        String nameWithExtension = XLSfile.getName();
        
        xlsRule = (XLSRule) drs.getMyRule("XLS File", order, nameWithExtension.substring(0, nameWithExtension.lastIndexOf('.')));
        
        for(Tab t : this.tabList){
            String chilName = rc.getChildrenName(fileName.replaceAll("\\s",""));
            
            String tabResult = t.convert(xlsRule, chilName, isSmartConvert);
            
            LinkedList<DoubleString> resultsByOrder
                                        = tabResults.containsKey(t.getOrder())?
                                          tabResults.get(t.getOrder()):
                                          new LinkedList<DoubleString>();
            
            resultsByOrder.addLast(new DoubleString(t.getName(),tabResult));
            
            tabResults.put(t.getOrder(), resultsByOrder);
        }
                
        //Overwritte standar conversion criteria
        if(xlsRule == null){
            xlsRule = new XLSRule();
        }
        
        String encoding = xlsRule.getEcoding();
        
        if(xlsRule.isSingleFrame()){
            String newName = rc.getNewName(fileName.replaceAll("\\s",""));
            
            newName = newName == null ? fileName.replaceAll("\\s","") : newName;
        
            String xmlContent = "<?xml version='1.0' encoding='"+encoding+"'?>"
                                + PropertiesManager.newLine + "<"+newName+">"+PropertiesManager.newLine;

            LinkedList<Integer> ordersInOrder = new LinkedList<Integer>();
            ordersInOrder.addAll(tabResults.keySet());
            
            java.util.Collections.sort(ordersInOrder);

            for(Integer tabOrder : ordersInOrder){
                for(DoubleString tr : tabResults.get(tabOrder)){
                    xmlContent += tr.xmlContent;
                }
            }

            xmlContent += "</"+newName+">";

            result.add(createXML(xmlContent,outputDirectory,fileName,encoding));
        }
        else{
            for(LinkedList<DoubleString> trList : tabResults.values()){
                for(DoubleString tr : trList){
                    result.add(createXML("<?xml version='1.0' encoding='"+encoding+"'?>"
                                         +PropertiesManager.newLine +tr.xmlContent,
                               outputDirectory,fileName+" - "+tr.name,encoding));
                }
            }
        }
        
        return result;
    }
    
    private BufferedWriter createXML(String xmlContent, String outputDirectory, String fileName, String encoding) throws Exception{
        Charset encodingToUse = java.nio.charset.Charset.availableCharsets().get(encoding);
        BufferedWriter xml = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(outputDirectory+"\\"+fileName+".xml"), encodingToUse));
        xml.write(xmlContent);
        xml.close();
        IConsoleService  ics = Modules.getConsoleService();
        ics.succes("File Converted: "+outputDirectory+"\\"+fileName+".xml");
        return xml;
    }

    String convertChildTab(String tab, boolean isSmartConvert) throws Exception {
        RenameCollection rc = RenameCollection.getInstance();
        String fileName = XLSfile.getName().replaceFirst("[.][^.]+$", "");
        String chilName = rc.getChildrenName(fileName.replace("\\s",""));
        return tabMap.get(tab).convert(xlsRule, chilName, isSmartConvert);
    }

    void clear() {
        xlsRule = null;
        tabList = new LinkedList<Tab>();
        tabMap = new HashMap<String,Tab>();
        
        for (int k = 0; k < sheets.getNumberOfSheets(); k++){
            Sheet s = sheets.getSheetAt(k);
            Tab t = new Tab(this,s,k);
            tabList.add(t);
            tabMap.put(s.getSheetName(),t);
        }
    }
}
