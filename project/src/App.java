import java.io.*;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import java.awt.*;
import java.awt.event.*;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JTable;
import javax.swing.table.JTableHeader;

import java.io.*;
import java.awt.Color;

class excel {
    private int sum;
    private static float students;

    public float getAverage(String pathname) {
        try {
            FileInputStream file = new FileInputStream(pathname);
            XSSFWorkbook workBook = new XSSFWorkbook(file);
            XSSFSheet sheet = workBook.getSheetAt(0); // Get Your Sheet.
            students = 0;
            for (Row row : sheet) {// For each Row.
                Cell cell = row.getCell(2); // Get the Cell at the Index / Column you want.

                switch (cell.getCellTypeEnum()) {
                    case NUMERIC:
                        sum += cell.getNumericCellValue();
                        students += 1;
                        // System.out.print(cell.getNumericCellValue() + "\t");
                        break;
                    case STRING:
                        // System.out.print(cell.getStringCellValue() + "\t");
                        break;
                    case BLANK:
                        System.out.print("blnk");
                        break;
                    case _NONE:
                        System.out.print("None");
                        break;
                    case BOOLEAN:
                        System.out.print("Bool");
                        break;
                    case ERROR:
                        System.out.print("Error");
                        break;
                    case FORMULA:
                        System.out.print("Form");
                        break;
                }
                System.out.print(cell + " ");
            }
            workBook.close();
            file.close();
            System.out.println(sum + " Average is : " + sum / students);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return sum / students;
    }
}
// Main class
public class App {
    JFrame win;
    JButton  btnavg, btnwrite, open;
    Label lbl1, lbl2, lbl3, result;
    TextField pathinput, input1, input2, input3, input4;
    JTable table;

    App() {
        win = new JFrame("Read/Write Excel File");
        lbl1 = new Label("Enter Path of your excel file: ");
        lbl1.setPreferredSize(new Dimension(300, 50));
        pathinput = new TextField("Path here");
        pathinput.setPreferredSize(new Dimension(300, 50));
        open = new JButton("Open");
        open.setPreferredSize(new Dimension(100, 50));
        btnavg = new JButton("Average Marks");
        btnavg.setPreferredSize(new Dimension(100, 50));
        result = new Label("Result ");
        result.setPreferredSize(new Dimension(100, 50));
        input1 = new TextField("S.No");
        input1.setPreferredSize(new Dimension(50, 50));
        input2 = new TextField("Name");
        input2.setPreferredSize(new Dimension(50, 50));
        input3 = new TextField("Marks");
        input3.setPreferredSize(new Dimension(50, 50));
        input4 = new TextField("Pass/Fail");
        input4.setPreferredSize(new Dimension(50, 50));
        btnwrite = new JButton("Write Data");
        btnwrite.setPreferredSize(new Dimension(50, 50));

        win.add(lbl1);
        win.add(pathinput);
        win.add(open);
        win.add(btnavg);
        win.add(result);
        win.add(input1);
        win.add(input2);
        win.add(input3);
        win.add(input4);
        win.add(btnwrite);

        // win.setLayout(new FlowLayout(FlowLayout.CENTER));
        win.setLayout(new GridLayout(3,4));
        win.setSize(600, 250);
        win.setVisible(true);
        win.addWindowListener(new WindowAdapter() {
            public void windowClosing(WindowEvent e) {
                win.dispose();
            }
        });
        open.addActionListener(new ButtonClickListener(win, pathinput, result, table, input1, input2, input3, input4));
        btnavg.addActionListener(
                new ButtonClickListener(win, pathinput, result, table, input1, input2, input3, input4));
        btnwrite.addActionListener(
                new ButtonClickListener(win, pathinput, result, table, input1, input2, input3, input4));
    }
    // Main driver method
    public static void main(String[] args) {
        App obj = new App();
        obj.getClass();
    }
}

class ButtonClickListener extends Exception implements ActionListener {
    JFrame win;
    TextField pathinput, input1, input2, input3, input4;
    Label result;
    JTable table;

    ButtonClickListener(JFrame win, TextField pathinput, Label result, JTable table, TextField input1, TextField input2,
            TextField input3, TextField input4) {
        this.win = win;
        this.pathinput = pathinput;
        this.result = result;
        this.table = table;
        this.input1 = input1;
        this.input2 = input2;
        this.input3 = input3;
        this.input4 = input4;
    }

    public void actionPerformed(ActionEvent e) {
        String command = e.getActionCommand();
        String filepath = pathinput.getText();

        if (command.equals("Open")) {
            try {
                File u = new File(filepath);
                Desktop d = Desktop.getDesktop();
                d.open(u);
            } catch (Exception evt) {
            }
        } else if (command.equals("Average Marks")) {
            try {
                excel obj = new excel();
                String avg = String.valueOf((obj.getAverage(filepath)));
                result.setText("Average is : " + avg);
            } catch (Exception evt) {
            }
        } else if (command.equals("Write Data")) {
            try {
                FileInputStream file = new FileInputStream(filepath);
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                XSSFSheet sheet = workbook.getSheetAt(0);
                Row row;
                Map<String, Object[]> studentData = new TreeMap<String, Object[]>();
                studentData.put("1", new Object[] { input1.getText(), input2.getText(), "78", "Pass" });
                Set<String> keyid = studentData.keySet();
                System.out.println("\n" + sheet.getLastRowNum() + " and " + sheet.getPhysicalNumberOfRows());
                int rowid = sheet.getPhysicalNumberOfRows()-1;
                for (String key : keyid) {

                    row = sheet.createRow(rowid++);
                    Object[] objectArr = studentData.get(key);
                    int cellid = 0;

                    for (Object obj : objectArr) {
                        Cell cell = row.createCell(cellid++);
                        cell.setCellValue((String) obj);
                        System.out.println(cell);
                    }
                }
                OutputStream os = new FileOutputStream(filepath);
                workbook.write(os);
                file.close();
                workbook.close();
            } catch (Exception exp) {
            }
        }
    }
}