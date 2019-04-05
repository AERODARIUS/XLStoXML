/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic.data;

import java.util.HashMap;
import java.util.LinkedList;
import xlstoxml.databackup.PropertiesManager;
import xlstoxml.service.Modules;
import xlstoxml.service.logic.rules.RenameCollection;

/**
 *
 * @author ASUS
 */
public class CREntity {
    private Tab parentTab;
    private LinkedList<Integer> keys = new LinkedList<Integer>();
    private HashMap<Integer,String> cells = new HashMap<Integer,String>();
    private HashMap<Integer,CRHeader> headers = new HashMap<Integer,CRHeader>();
    private boolean isConvertible = false;
    
    public CREntity(Tab t, HashMap<Integer, CRHeader> h){
        headers = h;
        parentTab = t;
    }
    
    public CREntity(Tab t){
        headers = null;
        parentTab = t;
    }

    void addCell(Integer k, String content){
        keys.add(k);
        cells.put(k,content);
        if(content != null && !content.isEmpty()){
            isConvertible = true;
        }
    }

    String convert(String tagName, String nameFromParent, boolean isSmartConvert) throws Exception {
        Modules.getConsoleService().log("Convert CREntity");
        String attributes = "";
        String content = "";
        
        if(headers != null){
            for(Integer k : keys){
                String cellContent = cells.get(k);
                    CRHeader crh = headers.get(k);

                    switch(crh.getType()){
                        case "Text":
                            content += "\t"+cellContent + PropertiesManager.newLine;
                        break;
                        case "CDATA":
                            content += "\t<![CDATA["+cellContent+"]]>" + PropertiesManager.newLine;
                        break;
                        case "Property":
                            attributes += crh.title+"='"+cellContent+"' ";
                        break;
                        case "Reference":
                            for(String tab : cellContent.split(PropertiesManager.newLine)){
                                content += "\t"+parentTab.convertBrotherTab(tab, isSmartConvert)
                                            + PropertiesManager.newLine;
                            }
                        break;
                        case "SubTag":
                            String cellTagName = crh.title.replaceAll("\\s+","");
                            content += "\t<"+cellTagName+">" + PropertiesManager.newLine +
                                    cellContent+PropertiesManager.newLine
                                    + "</"+cellTagName+">" + PropertiesManager.newLine;
                        break;
                        default:
                            throw new Exception("Invalid Column/Row Type in header");
                    }
            }
        }
        else if(!cells.isEmpty()){
            for(Integer k : keys){
                content += "\t"+cells.get(k) + PropertiesManager.newLine;
            }
        }
        
        RenameCollection rc = RenameCollection.getInstance();
        String newName = rc.getNewName(tagName+"-item");
           
        if(newName == null){
            newName = nameFromParent;
        }

        if(newName == null){
            newName = tagName+"-item";
        }
        
        newName = newName.replaceAll("\\s+", "");
        
        String result = "<"+newName+" "+attributes+">" + PropertiesManager.newLine;
        result += content;
        result += "</"+newName+">" + PropertiesManager.newLine;
        
        return result;
    }

    boolean isConvertible() {
        return isConvertible;
    }
}
