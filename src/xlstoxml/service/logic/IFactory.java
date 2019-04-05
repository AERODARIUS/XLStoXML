/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic;

import xlstoxml.service.logic.data.IXLSCollection;

/**
 *
 * @author ASUS
 */
public interface IFactory {
    public IXLSCollection getXLSCollection();
    
    public IRulesCollection getRulesCollection();
    
    public IRenameCollection getRenameCollection();
}
