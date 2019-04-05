/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic.data;

import xlstoxml.service.logic.IRulesCollection;
import xlstoxml.service.logic.IFactory;
import xlstoxml.service.logic.IRenameCollection;
import xlstoxml.service.logic.rules.RenameCollection;
import xlstoxml.service.logic.rules.RulesCollection;


/**
 *
 * @author ASUS
 */
public class Factory implements IFactory{

    private static IXLSCollection xlsCollection;
    
    private static IRulesCollection rulesCollection;
    
    private static IRenameCollection renameCollection;

    @Override
    public IXLSCollection getXLSCollection() {
        xlsCollection = XLSCollection.getInstance();
        return xlsCollection;
    }

    @Override
    public IRulesCollection getRulesCollection() {
        rulesCollection = RulesCollection.getInstance();
        return rulesCollection;
    }

    @Override
    public IRenameCollection getRenameCollection() {
        renameCollection = RenameCollection.getInstance();
        return renameCollection;
    }
}
