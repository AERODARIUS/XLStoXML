/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic;

/**
 *
 * @author DCruz
 */
public interface IRenameCollection {

    public void add(String originalName, String newName, Boolean renameChildren) throws Exception;

    public void update(int ind, String originalName, String newName, Boolean renameChildren) throws Exception;

    public void remove(int ind);

    public void move(int ind, int direction);
    
}
