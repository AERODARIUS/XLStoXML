/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic;

import java.io.BufferedWriter;
import java.io.File;
import java.util.HashSet;
import java.util.LinkedList;
import xlstoxml.service.IUploadNConvertService;
import xlstoxml.service.logic.data.Factory;
import xlstoxml.service.logic.data.IXLSCollection;

/**
 *
 * @author DCruz
 */
public class UploadNConvertService implements IUploadNConvertService{
    
    private UploadNConvertService() {
    }
    
    public static UploadNConvertService getInstance() {
        return UploadNConvertServiceHolder.INSTANCE;
    }
    
    private static class UploadNConvertServiceHolder {

        private static final UploadNConvertService INSTANCE = new UploadNConvertService();
    }

    @Override
    public void addFiles(File[] XLSfile) throws Exception {
        IFactory fy = new Factory();
        IXLSCollection xlsc = fy.getXLSCollection();
        xlsc.addAll(XLSfile);
    }

    @Override
    public void removeFiles(int[] indices) {
        IFactory fy = new Factory();
        IXLSCollection xlsc = fy.getXLSCollection();
        xlsc.removeAll(indices);
    }

    @Override
    public LinkedList<BufferedWriter> convert(int[] indices, String outputDirectory, boolean isSmartConvert) throws Exception {
        IFactory fy = new Factory();
        IXLSCollection xlsc = fy.getXLSCollection();
        LinkedList<BufferedWriter> result  = xlsc.convert(indices, outputDirectory, isSmartConvert);
        return result;
    }

    @Override
    public LinkedList<BufferedWriter> convertAll(String outputDirectory, boolean isSmartConvert) throws Exception {
        IFactory fy = new Factory();
        IXLSCollection xlsc = fy.getXLSCollection();
        LinkedList<BufferedWriter> result = xlsc.convertAll(outputDirectory, isSmartConvert);
        return result;
    }

    @Override
    public HashSet<String> updateFileList() throws Exception {
        IFactory fy = new Factory();
        IXLSCollection xlsc = fy.getXLSCollection();
        return xlsc.updateXLSList();
    }
}
