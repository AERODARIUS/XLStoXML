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
public class TagRename implements java.io.Serializable {
    public String oldName = "";
    public String newName = "";
    public Boolean renameChildren = false;
    public LinkedList<LinkedList<String>> renameMapping = new LinkedList<LinkedList<String>>();
}
