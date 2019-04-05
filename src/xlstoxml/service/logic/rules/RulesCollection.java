/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic.rules;

import xlstoxml.service.logic.IRulesCollection;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;
import xlstoxml.service.IConsoleService;
import xlstoxml.service.Modules;

public class RulesCollection implements IRulesCollection {
    private List<Rule> rules = new LinkedList<Rule>();
    private HashMap<String, rulesDictionary> rulesByType = new HashMap<String, rulesDictionary>();
    
    private static RulesCollection rc = null;
    
    private RulesCollection(){
        rulesByType.put("XLS File", new rulesDictionary());
        rulesByType.put("XLS Tab", new rulesDictionary());
        rulesByType.put("Column/Row", new rulesDictionary());
    }
    
    public static synchronized RulesCollection getInstance() {
       if (rc == null){
           rc = new RulesCollection();
       }
       return rc;
    }
    
    private class rulesDictionary{
        //Rule Name --> Rule Order --> Rule Id
        private HashMap<String,HashMap<Integer,Rule>> rulesWithOrderOrName = new HashMap<String,HashMap<Integer,Rule>>();
        
        public void add(Rule r) throws Exception{
            String ruleName = r.getApplyRuleToName();
            Integer ruleOrder = r.getApplyRuleToIndex();
            
            
            //Can be empty/0 but not null
            
            if(ruleName == null){
                throw new Error("Rule Name is null");
            }
            
            if(ruleOrder == null){
                throw new Error("Rule Order is null");
            }
            
            HashMap<Integer,Rule> rulesByOrder
                    = rulesWithOrderOrName.containsKey(ruleName)?
                        rulesWithOrderOrName.get(ruleName):
                        new HashMap<Integer,Rule>();
            
            
            if(rulesByOrder.containsKey(ruleOrder)){
                throw new Exception("Error: The rule already exist!");
            }
            else{
                rulesByOrder.put(ruleOrder, r);
            }
            
            rulesWithOrderOrName.put(ruleName, rulesByOrder);
        }
        
        public void remove(Rule r){
            HashMap<Integer,Rule> rulesByOrder
                    = rulesWithOrderOrName.get(r.getApplyRuleToName());
            
            rulesByOrder.remove(r.getApplyRuleToIndex());
        }
        
        public Rule get(String ruleName, Integer ruleOrder){
            Rule mergedRule = null;
            
            Rule exactMatch = null;
            
            Rule matchByName = null;
            
            Rule matchByOrder = null;
            
            Rule globalDefault = null;
            
            if(rulesWithOrderOrName.containsKey(ruleName)){
                HashMap<Integer,Rule> rulesByOrder
                        = rulesWithOrderOrName.get(ruleName);
                
                if(rulesByOrder.containsKey(ruleOrder)){
                    exactMatch = rulesByOrder.get(ruleOrder);
                }
                
                if(rulesByOrder.containsKey(0)){
                    matchByName = rulesByOrder.get(0);
                }
            }
            
            if(rulesWithOrderOrName.containsKey("")){
                HashMap<Integer,Rule> rulesByOrder
                        = rulesWithOrderOrName.get("");
                
                if(rulesByOrder.containsKey(ruleOrder)){
                    matchByOrder = rulesByOrder.get(ruleOrder);
                }
                
                if(rulesByOrder.containsKey(0)){
                    globalDefault = rulesByOrder.get(0);
                }
            }
            
            mergedRule = mergeRules(mergedRule, exactMatch);
            mergedRule = mergeRules(mergedRule, matchByName);
            mergedRule = mergeRules(mergedRule, matchByOrder);
            mergedRule = mergeRules(mergedRule, globalDefault);
                   
            return mergedRule;
        }

        private Rule mergeRules(Rule rule1, Rule rule2) {
            Rule mergedRule = null;
            
            if(rule1 == null){
                mergedRule = rule2;
            }
            else if(rule2 == null){
                mergedRule = rule1;
            }
            else{
                mergedRule = rule1.merge(rule2);
            }
            
            return mergedRule;
        }
    }
    
    @Override
    public List<Rule> getByIndex(int index) {
        List<Rule> result = new LinkedList<Rule>();
        Iterator<Rule> it = rules.iterator();
        Rule r = null;
        
        while(it.hasNext()){
            r = it.next();
            if(r.getApplyRuleToIndex() == index){
                result.add(r);
            }
        }
        
        return result;
    }
    
    @Override
    public List<Rule> getByName(String name) {
        List<Rule> result = new LinkedList<Rule>();
        Iterator<Rule> it = rules.iterator();
        Rule r = null;
        
        while(it.hasNext()){
            r = it.next();
            if(name.equals(r.getApplyRuleToName())){
                result.add(r);
            }
        }
        
        return result;
    }
    
    @Override
    public List<Rule> getByRule(String rule) {
        List<Rule> result = new LinkedList<Rule>();
        Iterator<Rule> it = rules.iterator();
        Rule r = null;
        
        while(it.hasNext()){
            r = it.next();
            if(rule.equals(r.getApplyRuleToName())){
                result.add(r);
            }
        }
        
        return result;
    }
    
    @Override
    public Rule getByPosition(int index) {
        return rules.get(index);
    }
    
    @Override
    public Rule get(int index, String name, String rule) {
        Iterator<Rule> it = rules.iterator();
        Rule r = null;
        
        while(it.hasNext() && r == null){
            r = it.next();
            if(!name.equals(r.getApplyRuleToName()) || r.getApplyRuleToIndex() != index || !rule.equals(r.getApplyRuleToName())){
                r = null;
            }
        }
        
        return r;
    }

    @Override
    public void put(Rule r) throws Exception{
        Integer k = rules.size();
        rules.add(r);
        rulesDictionary rd = rulesByType.get(r.getRuleType());
        rd.add(r);
        //print();
    }

    @Override
    public void remove(Set<Integer> indexes){
        Integer ruleCount = 0;
        Iterator<Rule> it = rules.iterator();
        
        while(it.hasNext()){
            Rule r = it.next();
            
            if(indexes.contains(ruleCount)){
                rulesByType.get(r.getRuleType()).remove(r);
                it.remove();
            }
            
            ruleCount++;
        }
        
        //print();
    }
    
    public int size() {
        return rules.size();
    }

    @Override
    public void reposition(int [] indexes){
        Rule [] originalRules = rules.toArray(new Rule[rules.size()]);
        Rule [] rulesRepositioned = new Rule[rules.size()];
        
        for(int k = 0; k < indexes.length; k++){
            
            rulesRepositioned[k] = originalRules[indexes[k]];
        }
        
        rules = new LinkedList<Rule>();
        
        rules.addAll(Arrays.asList(rulesRepositioned));
        
        //print();
    }

    @Override
    public void replaceRule(int index, Rule rule) throws Exception{
        Rule removed = rules.remove(index);
        rulesByType.get(removed.getRuleType()).remove(removed);
        
        rules.add(index, rule);
        rulesByType.get(rule.getRuleType()).add(rule);
        
        //print();
    }

    @Override
    public Rule getMyRule(String ruleType, Integer order, String name) {
        rulesDictionary rd = rulesByType.get(ruleType);
        
        if(rd != null){
            //Exact match & Matches by name & Matches by order
            Rule mergedRule = rd.get(name, order);
            
            return mergedRule;
        }
                
        return null;
    }
    
    private void print(){
        IConsoleService console = Modules.getConsoleService();
        
        for(Rule r : rules){
            String messageToSend = "\n";
            
            if(XLSRule.class.isInstance(r)){
                messageToSend += "XLSRule";
            }
            
            if(TabRule.class.isInstance(r)){
                messageToSend += "TabRule";
            }
            
            if(CRRule.class.isInstance(r)){
                messageToSend += "CRRule";
            }
            
            messageToSend += "\t\t"+r.getApplyRuleToIndex();
            messageToSend += "\t\t"+r.getApplyRuleToName();
            
            if(XLSRule.class.isInstance(r)){
                XLSRule xlsr = (XLSRule)r;
                messageToSend += "\t\t"+xlsr.getEcoding() +
                                "\t\t"+"------" +
                                "\t\t"+xlsr.isSingleFrame() +
                                "\t\t"+xlsr.isHasHeaders() +
                                "\t\t"+"------" +
                                "\t\t"+xlsr.isMatchByColumn();
            }
            
            if(TabRule.class.isInstance(r)){
                TabRule tr = (TabRule)r;
                messageToSend += "\t\t"+"------" +
                                "\t\t"+tr.getNewOrder() +
                                "\t\t"+tr.getIsMainWrapper()+
                                "\t\t"+tr.getHasHeaders() +
                                "\t\t"+"------" +
                                "\t\t"+tr.getMatchByColumn();
            }
            
            if(CRRule.class.isInstance(r)){
                CRRule crr = (CRRule)r;
                messageToSend += "\t\t"+"------";
                messageToSend += "\t\t"+"------";
                messageToSend += "\t"+"------";
                messageToSend += "\t"+"------";
                messageToSend += "\t"+crr.getColumnType();
                messageToSend += "\t"+"------";
            }
            
            console.log(messageToSend);
            console.log("");
            console.log( "----------------------------------------------" +
                                "----------------------------------------------" +
                                "----------------------------------------------" +
                                "----------------------------------------------");
            console.log("");
        }
        console.log("");
        console.log( "==============================================" +
                            "==============================================" +
                            "==============================================" +
                            "==============================================");
        console.log("");
    }
}
