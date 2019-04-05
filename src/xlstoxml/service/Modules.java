/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service;

import xlstoxml.service.logic.ConsoleService;
import xlstoxml.service.logic.DefineRulesService;
import xlstoxml.service.logic.TagRenameService;
import xlstoxml.service.logic.UploadNConvertService;

/**
 *
 * @author ASUS
 */
public class Modules {
    public static IDefineRulesService getDefineRulesService(){
        return DefineRulesService.getInstance();
    }
    
    public static IConsoleService getConsoleService(){
        return ConsoleService.getInstance();
    }

    public static ITagRenameService getTagRenameService() {
        return TagRenameService.getInstance();
    }

    public static IUploadNConvertService getUploadNConvertService() {
        return UploadNConvertService.getInstance();
    }
}
