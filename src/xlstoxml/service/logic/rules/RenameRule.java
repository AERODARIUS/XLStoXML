/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic.rules;

/**
 *
 * @author DCruz
 */
public class RenameRule{
    private static Integer cantInstances = 0;
    private Integer id;
    public String originalName;
    public String newName;
    public Boolean renameChildren = false;
    
    public RenameRule(){
        id = cantInstances++;
    }
    
    public Integer getId(){
        return id;
    }
}
