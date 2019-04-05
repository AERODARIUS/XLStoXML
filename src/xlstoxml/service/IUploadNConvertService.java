/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service;

import java.io.BufferedWriter;
import java.io.File;
import java.util.HashSet;
import java.util.LinkedList;

/**
 *
 * @author ASUS
 */
public interface IUploadNConvertService {

    public void addFiles(File[] XLSfile) throws Exception;

    public void removeFiles(int[] indices);

    public LinkedList<BufferedWriter> convert(int[] indices, String outputDirectory, boolean isSmartConvert) throws Exception;

    public LinkedList<BufferedWriter>  convertAll(String outputDirectory, boolean isSmartConvert) throws Exception;

    public HashSet<String> updateFileList() throws Exception;
    
}
