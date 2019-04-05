/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic;

import java.util.List;
import java.util.Set;
import xlstoxml.service.logic.rules.Rule;


/**
 *
 * @author ASUS
 */
public interface IRulesCollection{

    public List<Rule> getByIndex(int index);
    
    public List<Rule> getByName(String name);
    
    public List<Rule> getByRule(String name);
    
    public Rule getByPosition(int index);
    
    public Rule get(int index, String name, String rule);

    public void put(Rule r) throws Exception;

    public void remove(Set<Integer> indexes);
    
    public int size();

    public void reposition(int [] indexes);

    public void replaceRule(int index, Rule rule) throws Exception;

    public Rule getMyRule(String ruleType, Integer order, String name);
}
