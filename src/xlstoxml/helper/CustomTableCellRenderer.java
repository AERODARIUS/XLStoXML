/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.helper;

import java.awt.Color;
import java.awt.Component;
import java.util.HashMap;
import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;

/**
 *
 * @author ASUS
 */
public class CustomTableCellRenderer extends DefaultTableCellRenderer {
    private HashMap<Integer,HashMap<Integer,Boolean>> disabledCells;
    
    public CustomTableCellRenderer(HashMap<Integer,HashMap<Integer,Boolean>> dc){
        disabledCells = dc;
    }
    
    @Override
    public Component getTableCellRendererComponent(JTable table, Object value, boolean   isSelected, boolean hasFocus, int row, int column){ 
        Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
        
        if(!hasFocus){
            if(isSelected){
                c.setBackground(Color.BLUE);
            }
            else if(disabledCells.get(row).get(column)){
                c.setBackground(new java.awt.Color(220, 220, 220));//new java.awt.Color(255, 72, 72)
            }
            else{
                c.setBackground(Color.WHITE);
            }
        }
        else{
            c.setBackground(Color.DARK_GRAY);
        }
        return c;
    }
}
