/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic.rules;

import xlstoxml.service.Modules;

/**
 *
 * @author ASUS
 */
public class XLSRule extends Rule {
    private String ecoding = "UTF-8";
    private boolean singleFrame = false;
    private boolean matchByColumn = false; //otherwise match by rows
    private boolean hasHeaders = true;
    
    public XLSRule(){
        super();
    }

    public String getEcoding() {
        return ecoding;
    }

    public void setEcoding(String ecoding) {
        this.ecoding = ecoding;
    }

    public boolean isSingleFrame() {
        return singleFrame;
    }

    public void setSingleFrame(boolean singleFrame) {
        this.singleFrame = singleFrame;
    }

    public boolean isMatchByColumn() {
        return matchByColumn;
    }

    public void setMatchByColumn(boolean matchByColumn) {
        this.matchByColumn = matchByColumn;
    }

    public boolean isHasHeaders() {
        return hasHeaders;
    }

    public void setHasHeaders(boolean hasHeaders) {
        this.hasHeaders = hasHeaders;
    }

    @Override
    public String getRuleType() {
        return "XLS File";
    }

    @Override
    public Rule merge(Rule rule2) {
        XLSRule merged = new XLSRule();
        
        merged.ecoding = this.ecoding;
        merged.hasHeaders = this.hasHeaders;
        merged.matchByColumn = this.matchByColumn;
        merged.singleFrame = this.singleFrame;
        
        if("XLS File".equals(rule2.getRuleType())){
            XLSRule param = (XLSRule) rule2;
            
            if(merged.ecoding == null || merged.ecoding.equals("")){
                merged.ecoding = param.ecoding;
            }
        }
        
        return merged;
    }
}
