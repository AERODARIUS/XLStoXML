/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic.rules;

/**
 *
 * @author ASUS
 */
public class CRRule extends Rule {
    private String columnType;
    
    public CRRule(){
        super();
    }

    public String getColumnType() {
        return columnType;
    }

    public void setColumnType(String columnType) {
        this.columnType = columnType;
    }

    @Override
    public String getRuleType() {
        return "Column/Row";
    }

    @Override
    public Rule merge(Rule rule2) {
        CRRule merged = new CRRule();
        
        merged.columnType = this.columnType;
        
        if("Column/Row".equals(rule2.getRuleType())){
            CRRule param = (CRRule) rule2;
            
            if(merged.columnType == null || merged.columnType.equals("")){
                merged.columnType = param.columnType;
            }
        }
        
        return merged;
    }
}
