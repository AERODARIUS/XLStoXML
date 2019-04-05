/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service;

import javax.swing.JLabel;
import javax.swing.JTextPane;

/**
 *
 * @author ASUS
 */
public interface IConsoleService {
    
    public void succes(String msg);
    public void succes(String msg, JLabel label);
    public void log(String msg);
    public void log(String msg, JLabel label);
    public void warning(String msg);
    public void warning(String msg, JLabel label);
    public void error(String msg);
    public void error(String msg, JLabel label);

    public void addLogPane(JTextPane logPane);
    
    public String getMessagesToStore();

    public void clear();
}
