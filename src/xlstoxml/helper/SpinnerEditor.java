/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.helper;

import java.awt.Component;
import javax.swing.DefaultCellEditor;
import javax.swing.JSpinner;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.SpinnerNumberModel;

/**
 *
 * @author ASUS
 */
public class SpinnerEditor extends DefaultCellEditor {
    private JSpinner spinner;

    public SpinnerEditor()
    {
        super( new JTextField() );
        spinner = new JSpinner(new SpinnerNumberModel(0, 0, 100, 1));
        spinner.setBorder( null );
    }

    public Component getTableCellEditorComponent(
        JTable table, Object value, boolean isSelected, int row, int column)
    {
        spinner.setValue( value );
        return spinner;
    }

    public Object getCellEditorValue()
    {
        return spinner.getValue();
    }
}
