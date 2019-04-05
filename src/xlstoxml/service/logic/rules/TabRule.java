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
public class TabRule extends Rule {
    private Integer newOrder;
    private Boolean overWriteMatchBy;
    private Boolean overWriteHasHeaders;
    private Boolean IsMainWrapper;
    
    public TabRule(){
        super();
    }

    public Integer getNewOrder() {
        return newOrder;
    }

    public void setNewOrder(Integer newOrder) {
        this.newOrder = newOrder;
    }

    public void setNewOrder(Double newOrder) {
        this.newOrder = newOrder.intValue();
    }

    public Boolean getMatchByColumn() {
        return overWriteMatchBy;
    }

    public void setMatchByColumn(Boolean overWriteMatchBy) {
        this.overWriteMatchBy = overWriteMatchBy;
    }

    public Boolean getHasHeaders() {
        return overWriteHasHeaders;
    }

    public void setHasHeaders(Boolean overWriteHasHeaders) {
        this.overWriteHasHeaders = overWriteHasHeaders;
    }

    public void setIsMainWrapper(Boolean IsMainWrapper) {
        this.IsMainWrapper = IsMainWrapper;
    }

    public Boolean getIsMainWrapper() {
        return IsMainWrapper;
    }
    
    @Override
    public String getRuleType() {
        return "XLS Tab";
    }

    @Override
    public Rule merge(Rule rule2) {
        TabRule merged = new TabRule();
        
        merged.IsMainWrapper = this.IsMainWrapper;
        merged.newOrder = this.newOrder;
        merged.overWriteHasHeaders = this.overWriteHasHeaders;
        merged.overWriteMatchBy = this.overWriteMatchBy;
        
        if("XLS Tab".equals(rule2.getRuleType())){
            TabRule param = (TabRule) rule2;
            
            if(merged.newOrder == null || merged.newOrder == 0){
                merged.newOrder = param.newOrder;
            }
        }
        
        return merged;
    }
}
