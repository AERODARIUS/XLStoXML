/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic;

import java.util.HashSet;
import java.util.Set;
import java.util.logging.Level;
import java.util.logging.Logger;
import xlstoxml.service.IDefineRulesService;
import xlstoxml.service.logic.data.Factory;
import xlstoxml.service.logic.rules.CRRule;
import xlstoxml.service.logic.rules.Rule;
import xlstoxml.service.logic.rules.TabRule;
import xlstoxml.service.logic.rules.XLSRule;

/**
 *
 * @author ASUS
 */
public class DefineRulesService implements IDefineRulesService {
    
    private static DefineRulesService instance = new DefineRulesService();
    
    private DefineRulesService(){}

    public static IDefineRulesService getInstance(){
        return instance;
    }
  
    @Override
    public void setMatchByColumns(String option){
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void setSingleSheetType(boolean isSingleSheet) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void editFormatingRule() {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }

    @Override
    public void removeFormatingRule(int [] indexes) {
        IFactory fy = new Factory();
        
        IRulesCollection rules = fy.getRulesCollection();
        
        Set<Integer> rulesToRemove = new HashSet<Integer>();
        
        for(int k : indexes){
            rulesToRemove.add(k);
        }
        
        rules.remove(rulesToRemove);
    }

    @Override
    public String addFormatingRule(String matchBy, String enconding, Integer reorder, boolean hasHeaders, boolean isMainWrapper,
                                     String cellType, String applyRuleTo, Integer orderMatched, String nameMatched) {
        IFactory fy = new Factory();
        
        IRulesCollection rules = fy.getRulesCollection();
        
        Rule r = null;
        
        try {
            r = createNewRule(matchBy, enconding, reorder, hasHeaders, isMainWrapper,
                              cellType, applyRuleTo, orderMatched, nameMatched);
        } catch (Exception ex) {
            return "Error: Failed to add the rule. Check if you are not adding repeated rules!";
        }
        
        if(r != null){
            try {
                rules.put(r);
            } catch (Exception ex) {
                return "Error: Failed to add the rule. Check if you are not adding repeated rules!";
            }
        }
        else
            return "Error: Invalid option in 'Apply Rule To field'";
        
        return "Rule added successfully!";
    }

    @Override
    public void formatReposition(int[] indexes){
        IFactory fy = new Factory();
        
        IRulesCollection rules = fy.getRulesCollection();
        
        rules.reposition(indexes);
    }

    @Override
    public void setRuleValue(int ruleNumber, String field, String value, String [] allFields) throws Exception{
        IFactory fy = new Factory();
        
        IRulesCollection rules = fy.getRulesCollection();
        
        Boolean allFields5 = "yes".equals(allFields[5]);
        
        Boolean allFields6 = "yes".equals(allFields[6]);
        
        Rule newRule = createNewRule(allFields[8], allFields[3], Integer.parseInt(allFields[4]),
                                     allFields6, allFields5, allFields[7], allFields[0],
                                     Integer.parseInt(allFields[1]), allFields[2]);

        rules.replaceRule(ruleNumber, newRule);
    }

    @Override
    public Rule getMyRule(String ruleType, Integer order, String name) {
        IFactory fy = new Factory();
        
        IRulesCollection rules = fy.getRulesCollection();
        
        return rules.getMyRule(ruleType, order, name);
    }
    
    private Rule createNewRule(String matchBy, String enconding, Integer reorder, boolean hasHeaders, boolean isMainWrapper,
                                     String cellType, String applyRuleTo, Integer orderMatched, String nameMatched) throws Exception{
        Rule r = null;
        
        switch(applyRuleTo){
            case "XLS File":
                XLSRule xlsr = new XLSRule();
                xlsr.setEcoding(enconding);
                xlsr.setSingleFrame(isMainWrapper);
                xlsr.setHasHeaders(hasHeaders);
                xlsr.setMatchByColumn("columns".equals(matchBy));
                r = xlsr;
            break;
            case "XLS Tab":
                TabRule tabr = new TabRule();
                tabr.setNewOrder(reorder);
                tabr.setHasHeaders(hasHeaders);
                tabr.setIsMainWrapper(isMainWrapper);
                tabr.setMatchByColumn("columns".equals(matchBy));
                r = tabr;
            break;
            case "Column/Row":
                CRRule crr = new CRRule();
                crr.setColumnType(cellType);
                r = crr;
            break;
            default:
            break;
        }
        
        if(r != null){
            r.setApplyRuleToIndex(orderMatched);
            r.setApplyRuleToName(nameMatched);
        }
        else{
            Exception e = new Exception("Invalid option in 'Apply Rule To field'");
            throw e;
        }
        
        return r;
    }
}
