/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic.rules;

import java.util.HashMap;
import java.util.LinkedList;
import java.util.Objects;
import xlstoxml.service.IConsoleService;
import xlstoxml.service.logic.ConsoleService;
import xlstoxml.service.logic.IRenameCollection;

/**
 *
 * @author DCruz
 */
public class RenameCollection implements IRenameCollection{
    
    private LinkedList<RenameRule> renameRules = new LinkedList<RenameRule>();
    private HashMap<String,RenameRule> parents = new HashMap<String,RenameRule>();
    private HashMap<String,RenameRule> anytag = new HashMap<String,RenameRule>();
    
    private RenameCollection() {
    }
    
    public static RenameCollection getInstance() {
        return RenameCollectionHolder.INSTANCE;
    }

    public String getNewName(String fileName) {
        RenameRule rr = anytag.get(fileName);
        return rr == null || rr.newName == null ? null : rr.newName.replace("\\s","");
    }

    public String getChildrenName(String fileName) {
        RenameRule rr = parents.get(fileName);
        return rr == null || rr.newName == null ? null : rr.newName.replace("\\s","");
    }
    
    private static class RenameCollectionHolder {
        private static final RenameCollection INSTANCE = new RenameCollection();
    }
    
    @Override
    public void add(String originalName, String newName, Boolean renameChildren) throws Exception{
        if((renameChildren && !parents.containsKey(originalName))
            || (!renameChildren && !anytag.containsKey(originalName))){
            RenameRule r = new RenameRule();
            r.originalName = originalName.replace("\\s","").isEmpty() ? null : originalName.replace("\\s","");
            r.newName = newName.replace("\\s","").isEmpty() ? null : newName.replace("\\s","");
            r.renameChildren = renameChildren;
            renameRules.add(r);

            if(renameChildren){
                parents.put(originalName, r);
            }
            else{
                anytag.put(originalName, r);
            }
        }
        else{
            throw new Exception("Rename rule already added!");
        }
        
        //print();
    }
    
    @Override
    public void update(int ind, String originalName, String newName, Boolean renameChildren) throws Exception {
        RenameRule r = renameRules.get(ind);
        
        if((renameChildren && !parents.containsKey(originalName))
            || (!renameChildren && !anytag.containsKey(originalName))
            || (renameChildren.equals(r.renameChildren) && originalName.equals(r.originalName))){

            if(r.renameChildren){
                parents.remove(r.originalName);
            }
            else{
                anytag.remove(r.originalName);
            }

            r.originalName = originalName;
            r.newName = newName;
            r.renameChildren = renameChildren;

            if(r.renameChildren){
                parents.put(r.originalName, r);
            }
            else{
                anytag.put(r.originalName, r);
            }
        }
        else{
            throw new Exception("Update fail: Already exists a rule with this properties!");
        }
        
        //print();
    }

    @Override
    public void remove(int ind) {
        RenameRule r = renameRules.remove(ind);
        
        if(r.renameChildren){
            parents.remove(r.originalName);
        }
        else{
            anytag.remove(r.originalName);
        }
        
        //print();
    }

    @Override
    public void move(int ind, int direction) {
        RenameRule r1 = renameRules.get(ind);
        
        if(ind+direction < renameRules.size() && ind+direction > -1){
            RenameRule r2 = renameRules.get(ind+direction);
            String originalName = r1.originalName;
            String newName = r1.newName;
            Boolean renameChildren = r1.renameChildren;

            r1.originalName = r2.originalName;
            r1.newName = r2.newName;
            r1.renameChildren = r2.renameChildren;

            r2.originalName = originalName;
            r2.newName = newName;
            r2.renameChildren = renameChildren;
        }
        
        print();
    }
    
    private void print(){
        IConsoleService cs = ConsoleService.getInstance();
        for(RenameRule rr : renameRules){
            String msg = "";
            msg += rr.originalName+"\t\t\t";
            msg += rr.newName+"\t\t\t";
            msg += rr.renameChildren;
            cs.log(msg);
        }
        cs.log("------------------------------------------------------"+
                      "------------------------------------------------------");
        cs.log("");
        cs.log("");
    }
}
