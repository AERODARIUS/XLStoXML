/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package xlstoxml.helper;

import java.awt.Color;
import java.awt.LayoutManager;
import java.io.BufferedWriter;
import java.util.LinkedList;
import javax.swing.JList;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import xlstoxml.service.IConsoleService;
import xlstoxml.service.IUploadNConvertService;

/**
 *
 * @author DCruz
 */
public class ConvertThread implements Runnable {
   private Thread t;
   private IConsoleService console;
   private IUploadNConvertService converter;
   private JList list;
   private JLabel label;
   private LayoutManager layout;
   private JLabel convertMessages;
   private JLabel gearsLabel;
   private JPanel jPanel10;
   private JScrollPane jScrollPane1;
   private boolean isSmartConvert;
   
   public ConvertThread(boolean isSmartConvert, IConsoleService console, IUploadNConvertService converter,
                 javax.swing.JList jList1, LayoutManager prevLayout, JLabel jLabel21,
                 JLabel uploadNConvertMessages, JLabel gears, JPanel jPanel10, JScrollPane jScrollPane1){
       this.console = console;
       this.converter = converter;
       this.list = jList1;
       this.label = jLabel21;
       this.layout = prevLayout;
       this.convertMessages = uploadNConvertMessages;
       this.gearsLabel = gears;
       this.jPanel10 = jPanel10;
       this.jScrollPane1 = jScrollPane1;
       this.isSmartConvert = isSmartConvert;
   }
   
    private void hideOverlay(LayoutManager layoutManager){
        if(layout != null){
            jPanel10.setLayout(layoutManager);
        }
        
        gearsLabel.setVisible(false);
        jScrollPane1.setVisible(true);
    }
   
   public void run(){
        int[] indices = list.getSelectedIndices();

        try{
            LinkedList<BufferedWriter> conversions;
                    
            if(indices.length > 0){
                conversions = converter.convert(indices, label.getText(), isSmartConvert);
            }
            else{
                conversions = converter.convertAll(label.getText(), isSmartConvert);
            }
            
            console.log("Successful conversions:  " + conversions.size(), convertMessages);
            
            if(conversions.size() < indices.length){
                list.setBackground(Color.RED);
                console.error("Not all conversions succeeded.");
            }
            else{
                list.setBackground(Color.GREEN);
            }
            
            //Successful conversions:
        } catch (Exception ex) {
            list.setBackground(Color.RED);
            console.error("Conversion fail!");
            console.error(ex.getMessage());
        }
        
       hideOverlay(layout);
    }

   public void start(){
        if (t == null){
            t = new Thread(this, "convertXLStoXML");
            t.start();
        }
   }

    public void waitUntilFinish() throws Exception {
        t.join();
    }
}
