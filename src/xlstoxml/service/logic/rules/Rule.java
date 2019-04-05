package xlstoxml.service.logic.rules;

/**
 *
 * @author ASUS
 */
public abstract class Rule {
    private String applyRuleToName = null;
    private Integer applyRuleToIndex;
    private static Integer instancesCount = 0;
    private Integer id;
    
    public Rule(){
        id = instancesCount++;
    }

    public String getApplyRuleToName() {
        return applyRuleToName;
    }

    public void setApplyRuleToName(String applyRuleToName) {
        this.applyRuleToName = applyRuleToName;
    }

    public Integer getApplyRuleToIndex() {
        return applyRuleToIndex;
    }

    public void setApplyRuleToIndex(Integer applyRuleToIndex) {
        this.applyRuleToIndex = applyRuleToIndex;
    }
    
    public void setApplyRuleToIndex(Double applyRuleToIndex) {
        this.applyRuleToIndex = applyRuleToIndex.intValue();
    }

    public abstract String getRuleType();
    
    public Integer getId(){
        return id;
    }

    public abstract Rule merge(Rule rule2);
}
