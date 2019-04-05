/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service;

import xlstoxml.service.logic.rules.Rule;

/**
 *
 * @author ASUS
 */
public interface IDefineRulesService {
    
    public void setMatchByColumns(String option);// OK
    
    public void setSingleSheetType(boolean isSingleSheet);// OK
    
    public void editFormatingRule();
    
    public void removeFormatingRule(int [] indexes);

    public String addFormatingRule(String matchBy,
                                      String enconding,
                                      Integer reorder,
                                      boolean hasHeaders,
                                      boolean isMainWrapper,
                                      String cellType,
                                      String applyRuleTo,
                                      Integer orderMatched,
                                      String nameMatched);

    public void formatReposition(int[] indexes);

    public void setRuleValue(int ruleNumber, String field, String value, String [] allFields) throws Exception;

    public Rule getMyRule(String ruleType, Integer order, String name);
}
