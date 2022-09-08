import java.awt.*;
import javax.swing.*;
import java.awt.event.*;
import java.io.*;
import java.net.*;
  
class awtdesk extends JFrame implements ActionListener {
  
    // frame
    static JFrame f;
      
    // Main Method 
    public static void main(String args[])
    {
        awtdesk d = new awtdesk();
  
        // create a frame
        f = new JFrame("desktop");
  
        // create a panel
        JPanel p = new JPanel();
  
        // create a button
        JButton b = new JButton("launch");
  
        // add Action Listener
        b.addActionListener(d);
  
        p.add(b);
        f.add(p);
        f.show();
        f.setSize(200, 200);
    }
  
    // if button is pressed
    public void actionPerformed(ActionEvent e)
    {
        try {
  
            File u = new File("data.xlsx");
  
            Desktop d = Desktop.getDesktop();
            d.open(u);
            
        }
        catch (Exception evt) {
        }
    }
}