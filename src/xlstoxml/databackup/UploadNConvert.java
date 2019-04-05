/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.databackup;

import java.util.LinkedList;

/**
 *
 * @author DCruz
 */
public class UploadNConvert implements java.io.Serializable{
    public String outputFolder = "";
    public LinkedList<String> xlsList = new LinkedList<String>();
}
