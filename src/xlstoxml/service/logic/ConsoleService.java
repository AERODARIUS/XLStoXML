/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.service.logic;

import java.awt.Color;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JLabel;
import javax.swing.JTextPane;
import javax.swing.text.BadLocationException;
import javax.swing.text.PlainDocument;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;
import javax.swing.text.StyledDocument;
import xlstoxml.databackup.PropertiesManager;
import xlstoxml.XLStoXML;
import xlstoxml.service.IConsoleService;

/**
 *
 * @author ASUS
 */
public class ConsoleService implements IConsoleService{
    
    private static ConsoleService instance = new ConsoleService();
    
    private ConsoleService(){}

    public static IConsoleService getInstance(){
        return instance;
    }
    
    private static JTextPane log;
    private String messagesToStore = "";

    @Override
    public void addLogPane(JTextPane logPane) {
        log = logPane;
        log("Session Started.");
    }
    
    @Override
    public void log(String msg){
        SimpleAttributeSet keyWord = createSimpleAttributeSet(Color.LIGHT_GRAY,null);
        this.addGenericMessage(msg, keyWord);
    }

    @Override
    public void warning(String msg) {
        SimpleAttributeSet keyWord = createSimpleAttributeSet(Color.YELLOW,null);
        this.addGenericMessage(msg, keyWord);
    }

    @Override
    public void error(String msg) {
        SimpleAttributeSet keyWord = createSimpleAttributeSet(Color.RED,null);
        this.addGenericMessage(msg, keyWord);
    }

    @Override
    public void succes(String msg) {
        SimpleAttributeSet keyWord = createSimpleAttributeSet(Color.GREEN,null);
        this.addGenericMessage(msg, keyWord);
    }
    
    private void addGenericMessage(String msg, SimpleAttributeSet keyWord){
        try {
            if(log != null){
                StyledDocument doc = log.getStyledDocument();
                String consoleMsg = "[" + new java.util.Date()+"] ";
                consoleMsg += msg+"\n\n\n";
                doc.insertString(doc.getLength(), consoleMsg, keyWord);
                messagesToStore += consoleMsg.replaceAll("\n", PropertiesManager.newLine);
            }
            else{
                System.out.println(msg);
            }
        } catch (Exception ex) {
            Logger.getLogger(XLStoXML.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    private SimpleAttributeSet createSimpleAttributeSet(Color c1, Color c2){
            SimpleAttributeSet keyWord = new SimpleAttributeSet();
            StyleConstants.setForeground(keyWord, c1);
            StyleConstants.setFontFamily(keyWord, "monospaced");
            
            if(c2 != null){
                StyleConstants.setBackground(keyWord, c2);
            }
            
            return keyWord;
    }

    @Override
    public void succes(String msg, JLabel label){
        this.succes(msg);
        label.setText(msg);
        label.setForeground(Color.decode("#006301"));
    }

    @Override
    public void log(String msg, JLabel label){
        this.log(msg);
        label.setText(msg);
        label.setForeground(Color.decode("#3366FF"));
    }

    @Override
    public void warning(String msg, JLabel label){
        this.warning(msg);
        label.setText(msg);
        label.setForeground(Color.decode("#FF6600"));
    }

    @Override
    public void error(String msg, JLabel label){
        this.error(msg);
        label.setText(msg);
        label.setForeground(Color.red);
    }

    @Override
    public String getMessagesToStore() {
        return messagesToStore;
    }

    @Override
    public void clear() {
        try {
            StyledDocument doc = log.getStyledDocument();
            doc.remove(0, doc.getLength());
        } catch (BadLocationException ex) {
            this.error("Error while cleaning console");
        }
    }
}
