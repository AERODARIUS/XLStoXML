/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service;

/**
 *
 * @author ASUS
 */
public interface ITagRenameService {

    public void addRule(String originalName, String newName, Boolean renameChildren) throws Exception;

    public void updateRule(Integer ind, String originalName, String newName, Boolean renameChildren) throws Exception;

    public void removeRule(Integer ind);

    public void moveDown(Integer ind);

    public void moveUp(Integer ind);
    
}
