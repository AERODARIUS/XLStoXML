/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic;

import xlstoxml.service.ITagRenameService;
import xlstoxml.service.logic.data.Factory;

/**
 *
 * @author DCruz
 */
public class TagRenameService implements ITagRenameService {
    
    private TagRenameService() {
    }
    
    public static TagRenameService getInstance() {
        return TagRenameServiceHolder.INSTANCE;
    }
    
    private static class TagRenameServiceHolder {

        private static final TagRenameService INSTANCE = new TagRenameService();
    }

    @Override
    public void addRule(String originalName, String newName, Boolean renameChildren) throws Exception {
        IFactory fy = new Factory();
        IRenameCollection rc = fy.getRenameCollection();
        rc.add(originalName, newName, renameChildren);
    }

    @Override
    public void updateRule(Integer ind, String originalName, String newName, Boolean renameChildren) throws Exception {
        IFactory fy = new Factory();
        IRenameCollection rc = fy.getRenameCollection();
        rc.update(ind, originalName, newName, renameChildren);
    }

    @Override
    public void removeRule(Integer ind) {
        IFactory fy = new Factory();
        IRenameCollection rc = fy.getRenameCollection();
        rc.remove(ind);
    }

    @Override
    public void moveDown(Integer ind) {
        IFactory fy = new Factory();
        IRenameCollection rc = fy.getRenameCollection();
        rc.move(ind,1);
    }

    @Override
    public void moveUp(Integer ind) {
        IFactory fy = new Factory();
        IRenameCollection rc = fy.getRenameCollection();
        rc.move(ind,-1);
    }
}
