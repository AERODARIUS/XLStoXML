/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic.data;

import java.io.File;
import java.util.HashMap;
import java.util.LinkedList;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import xlstoxml.databackup.PropertiesManager;
import xlstoxml.service.IDefineRulesService;
import xlstoxml.service.Modules;
import xlstoxml.service.logic.rules.CRRule;
import xlstoxml.service.logic.rules.RenameCollection;
import xlstoxml.service.logic.rules.TabRule;
import xlstoxml.service.logic.rules.XLSRule;

/**
 *
 * @author ASUS
 */
public class Tab {
    private XLS parent;
    private String name;
    private Integer order;
    private LinkedList<CREntity> crList = new LinkedList<CREntity>();
    HashMap<Integer, CRHeader> headers = new HashMap<Integer, CRHeader>();
    private Sheet sheet;
    private final String emptyRegEx = "^\\s*$";
    //To prevent cylcle dependencies
    private Boolean convertStarted = false;
    
    private String convertResultCache = null;

    Tab(XLS xls, Sheet s, int o) {
        //Modules.getConsoleService().log("SHEET NAME: "+s.getSheetName());
        sheet = s;
        name = s.getSheetName();
        order = o;
        parent = xls;
    }

    public Integer getOrder() {
        return order;
    }

    public String getName() {
        return name;
    }
    
    String convert(XLSRule xlsRule, String nameFromParent, boolean isSmartConvert) throws Exception {
        Modules.getConsoleService().log("Convert Tab "+name);
        Modules.getConsoleService().log("convertStarted "+convertStarted);
        RenameCollection rc = RenameCollection.getInstance();
        
        if(convertStarted && convertResultCache == null){
            throw new Exception("Convert fails because of circular dependency between file tabs!");
        }
        else if(convertResultCache != null){
            return convertResultCache;
        }
        
        convertStarted = true;
        
        String result = "";
                
        IDefineRulesService drs = Modules.getDefineRulesService();
        TabRule tabRule = (TabRule) drs.getMyRule("XLS Tab", order+1, name);
        
        //Default convertions criteria
        Boolean hasHeader = true;
        Boolean matchByColumn = false;
        Boolean isMainWrapper = true;
        
        //Overwritte standar conversion criteria
        if(tabRule != null){
            int newOrder = tabRule.getNewOrder();
            
            if(newOrder > 0){
                order = newOrder;
            }
            
            hasHeader = tabRule.getHasHeaders();
            matchByColumn = tabRule.getMatchByColumn();
            isMainWrapper = tabRule.getIsMainWrapper();
        }
        else if(xlsRule != null){
            hasHeader = xlsRule.isHasHeaders();
            matchByColumn = xlsRule.isMatchByColumn();
        }
        
        //First: Initialize to generate CREntity
        if(matchByColumn){
            parseByColumns(hasHeader, sheet);
        }
        else{
            parseByRows(hasHeader, sheet);
        }
        
        String chilName = rc.getChildrenName(sheet.getSheetName().replaceAll("\\s+",""));
        Modules.getConsoleService().log("Convert child:"+chilName);
        
        //Second: Invoke CREntity.convert()isSmarftConvert
        for(CREntity cre : crList){
            if(!isSmartConvert || cre.isConvertible()){
                result += cre.convert(sheet.getSheetName().replaceAll("\\s+",""), chilName, isSmartConvert)
                          + PropertiesManager.newLine;
            }
        }
         
        String newName = rc.getNewName(sheet.getSheetName().replace("\\s",""));
            
        if(newName == null){
            newName = nameFromParent;
        }

        if(newName == null){
            newName = sheet.getSheetName().replace("\\s","");
        }
            
        if(isMainWrapper){
            convertResultCache = "<"+newName+">" + PropertiesManager.newLine
                                 + result +
                                 "</"+newName+">" + PropertiesManager.newLine;
        }
        else{
            convertResultCache = result;
        }
        
        return convertResultCache;
    }
    
    //Iterate over the rows
    private void parseByRows(Boolean hasHeader, Sheet sheet) throws Exception{
        int rows = sheet.getLastRowNum()+1;
        int row = 0;
        int col = -1;
        
        if(hasHeader){
            while(col == -1 && row < rows){
                col = getFirstCellNotEmpty(sheet.getRow(row++));
            }
            
            if(col != -1 && row < rows){
                //Headers
                Row cellsInHeader = sheet.getRow(row-1);
                int cellSize = cellsInHeader.getLastCellNum()+1;

                String content;

                for(int k = col; k < cellSize; k++){
                    Cell c = cellsInHeader.getCell(k);
                    
                    if(c!= null){
                        content = c.getStringCellValue();

                        if(!content.matches(emptyRegEx)){
                            CRHeader crh = createCRHeader(content, k);
                            headers.put(k,crh);
                        }
                    }
                }

                while(row < rows){
                    Row cells = sheet.getRow(row++);
                    if(cells != null){
                        CREntity cre = new CREntity(this,headers);

                        for(Integer k : headers.keySet()){
                            content = cells.getCell(k) != null ?
                                        cells.getCell(k).getStringCellValue() : "";
                            cre.addCell(k, content);
                        }

                        crList.add(cre);
                    }
                }
            }
        }
        else{
            while(row < rows){
                Row cells = sheet.getRow(row++);
                
                CREntity cre = new CREntity(this);
                
                for(int k = 0; k < cells.getLastCellNum()+1; k++){
                    cre.addCell(k, cells.getCell(k).getStringCellValue());
                }
                
                crList.add(cre);
            }
        }
    }
    
    //If all the cells are tempty, returns the last cell
    private int getFirstCellNotEmpty(Row r){
        int k = 0;
        int cellSize = r.getLastCellNum()+1;
        boolean isEmpty = true;
        
        while(k < cellSize && isEmpty){
            String content = r.getCell(k).getStringCellValue();
            isEmpty = content == null || content.matches(emptyRegEx);
            k++;
        }
        
        if(k >= cellSize){
            k = -1;
        }
        
        return k;
    }

    private int getFirstCellNotEmpty(int col, Sheet sheet) {
        int r = 0;
        int rows = sheet.getLastRowNum()+1;
        int cols = 0;
        boolean isEmpty = true;
        
        for(int k = 0; k < rows; k++){
            int caux = sheet.getRow(k).getLastCellNum()+1;
            cols = cols < caux ? caux : cols;
        }

        while(r < rows && isEmpty){
            Row row = sheet.getRow(r);
            
            if(col <= row.getRowNum()){
                String content = sheet.getRow(r).getCell(col).getStringCellValue();
                isEmpty = content == null || content.matches(emptyRegEx);
            }
            
            r++;
        }
        
        if(r >= cols){
            r = -1;
        }
        
        return r;
    }
    
    //Iterate over the columns
    private void parseByColumns(Boolean hasHeader, Sheet sheet) throws Exception{
        int cols = sheet.getRow(0).getLastCellNum();
        int rows = sheet.getLastRowNum()+1;
        int col = 0;
        int row = -1;
        
        if(hasHeader){
            while(row == -1 && col < cols){
                row = getFirstCellNotEmpty(col++, sheet);
            }

            if(row != -1 && col < cols){
                //Headers
                Row cellsInHeader = sheet.getRow(row-1);
                int cellSize = cellsInHeader.getLastCellNum()+1;

                String content;

                for(int k = col; k < cellSize; k++){
                    content = cellsInHeader.getCell(k).getStringCellValue();

                    if(!content.matches(emptyRegEx)){
                        CRHeader crh = createCRHeader(content, k);
                        headers.put(k,crh);
                    }
                }

                while(col < cols){
                    CREntity cre = new CREntity(this,headers);

                    Row cells = sheet.getRow(col++);

                    for(Integer k : headers.keySet()){
                        content = cells.getCell(k).getStringCellValue();
                        cre.addCell(k, content);
                    }

                    crList.add(cre);
                }
            }
        }
        else{
            while(col < cols){                
                CREntity cre = new CREntity(this);
                int k = 0;

                int colSize = 0;

                for(int r = 0; r < rows; r++){
                    int caux = sheet.getRow(r).getLastCellNum()+1;
                    colSize = colSize < caux ? caux : colSize;
                }
                
                for(int r = 0; r < rows; r++){
                    Row currRow = sheet.getRow(r);
                    
                    if(col <= currRow.getLastCellNum()){
                        cre.addCell(r, currRow.getCell(col).getStringCellValue());
                    }
                }
                
                crList.add(cre);
                col++;
            }
        }
    }
    
    //Primero buscar los headers y despues las celdas
    //Salntenado espacios

    private CRHeader createCRHeader(String title, Integer position) throws Exception{
        IDefineRulesService drs = Modules.getDefineRulesService();
        CRRule crRule = (CRRule) drs.getMyRule("Column/Row", position+1, title);
        
        CRHeader crh;
        
        if(crRule != null){
            switch(crRule.getColumnType()){
                case "Text":
                    crh = new Text();
                break;
                case "CDATA":
                    crh = new CDATA();
                break;
                case "Property":
                    crh = new Property();
                break;
                case "Reference":
                    crh = new Ref();
                break;
                case "SubTag":
                    crh = new SubTag();
                break;
                default:
                    throw new Exception("Invalid Column/Row Type in property");
            }
        }
        else{
            crh = new SubTag();
        }
        
        crh.title = title;
        crh.order = position;
        
        return crh;
    }

    String convertBrotherTab(String tab, boolean isSmartConvert) throws Exception {
        return parent.convertChildTab(tab, isSmartConvert);
    }
}
