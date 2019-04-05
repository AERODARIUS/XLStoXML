/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic.data;

/**
 *
 * @author ASUS
 */
public abstract class CRHeader {
    public String title;
    public int order;
    
    public abstract String getType();
}
