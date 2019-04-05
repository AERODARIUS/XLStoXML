/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic.data;

import java.io.BufferedWriter;
import java.io.File;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import org.apache.poi.ss.usermodel.Sheet;
import xlstoxml.service.IConsoleService;
import xlstoxml.service.Modules;

/**
 *
 * @author ASUS
 */
public class XLSCollection implements IXLSCollection{
    private LinkedList<XLS> xlsFiles = new LinkedList<XLS>();
    private HashSet<String> paths = new HashSet<String>();
    
    private static XLSCollection xlsc = null;
    
    private XLSCollection(){
        
    }
    
    public static synchronized XLSCollection getInstance() {
       if (xlsc == null){
           xlsc = new XLSCollection();
       }
       return xlsc;
    }

    @Override
    public void addAll(File[] XLSfiles) throws Exception {
        for(File XLSfile : XLSfiles){
            String absolutePath = XLSfile.getAbsolutePath();
            
            if(!paths.contains(absolutePath)){
                XLS xls = new XLS(XLSfile);
                xlsFiles.addLast(xls);
                paths.add(absolutePath);
            }
        }
        
        //print();
    }

    @Override
    public void removeAll(int[] indices) {
        LinkedList<XLS> newXlsFiles = new LinkedList<XLS>();
        HashSet<Integer> excludedIndices = arrayToList(indices);
        paths = new HashSet<String>();
        Integer count = 0;
        
        for(XLS xls : xlsFiles){
            if(!excludedIndices.contains(count)){
                newXlsFiles.addLast(xls);
                paths.add(xls.XLSfile.getAbsolutePath());
            }
            
            count++;
        }
        
        xlsFiles = newXlsFiles;
    
        //print();
    }

    @Override
    public void removeAll() {
        LinkedList<XLS> newXlsFiles = new LinkedList<XLS>();
        paths = new HashSet<String>();
        //print();
    }

    @Override
    public LinkedList<BufferedWriter> convert(int[] indices, String outputDirectory, boolean isSmartConvert)  throws Exception{
        LinkedList<BufferedWriter> xmlList = new LinkedList<BufferedWriter>();
        
        for(int k = 0;k < indices.length;k++){
            XLS xls = xlsFiles.get(indices[k]);
            xmlList.addAll(xls.convert(outputDirectory, indices[k]+1, isSmartConvert));
        }
        
        clearConvertData();
        
        return xmlList;
    }

    @Override
    public LinkedList<BufferedWriter> convertAll(String outputDirectory, boolean isSmartConvert) throws Exception{
        LinkedList<BufferedWriter> xmlList = new LinkedList<BufferedWriter>();
        int order = 1;
        
        for(XLS xls : xlsFiles){
            xmlList.addAll(xls.convert(outputDirectory, order++, isSmartConvert));
        }
        
        clearConvertData();
        
        return xmlList;
    }
    
    private void clearConvertData(){
        for(XLS xls : xlsFiles){
            xls.clear();
        }
    }
    
    private HashSet<Integer> arrayToList(int[] indices){
        HashSet<Integer> resIndices = new HashSet<Integer>();
        
        for(int k = 0; k < indices.length; k++){
            resIndices.add(indices[k]);
        }
        
        return resIndices;
    }

    @Override
    public HashSet<String> updateXLSList() throws Exception {
        HashSet<String> pathsList = new HashSet<String>();
        paths = new HashSet<String>();
        
        for (Iterator<XLS> it = xlsFiles.iterator(); it.hasNext();) {
            XLS xls = it.next();
            
            if(xls.XLSfile.exists() && !xls.XLSfile.isDirectory()) { 
                pathsList.add(xls.XLSfile.getAbsolutePath());
                xls.reFresh();
            }
            else{
                it.remove();
            }
        }
        
        print();
        
        return pathsList;
    }
    
    private void print(){
        String strToPrint = "";
                
        IConsoleService console = Modules.getConsoleService();
        
        for(XLS x : xlsFiles){
            strToPrint += "\n"+x.XLSfile.getAbsolutePath();
            
            String sheetsNames = "";
            
            for (int k = 0; k < x.sheets.getNumberOfSheets(); k++){
                Sheet s = x.sheets.getSheetAt(k);
                sheetsNames += "\n\t"+s.getSheetName();
            }
            
            strToPrint += sheetsNames;
        }
        
        strToPrint += "\n----------------------------------------------"
                    + "----------------------------------------------"
                    + "----------------------------------------------";
        
        console.log(strToPrint);
    }
}
