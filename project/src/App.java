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

class excel {
    private int sum;
    private String pathname = "data.xlsx";
    private static float students;

    public void writedata() {
        System.out.println("WriteData() function exceuting");
        try {
            FileInputStream file = new FileInputStream("data2.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            // Iterator<Row> rowIterator = sheet.iterator();
            Row row;
            Map<String, Object[]> studentData = new TreeMap<String, Object[]>();
            studentData.put("1", new Object[] { "1", "newguy", "78", "Pass" });
            Set<String> keyid = studentData.keySet();
            System.out.println("\n"+sheet.getLastRowNum() + " and " + sheet.getPhysicalNumberOfRows());
            int rowid = sheet.getPhysicalNumberOfRows();
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
            OutputStream os = new FileOutputStream("data2.xlsx");
            workbook.write(os);
            file.close();
            workbook.close();
            // Desktop.getDesktop().open(new File("data.xlsx"));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void getAverage() {
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
    }
    public void getData() {
        // Try block to check for exceptions
        try {
            // Reading file from local directory
            FileInputStream file = new FileInputStream(pathname);

            // Create Workbook instance holding reference to
            // .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();

            // Till there is an element condition holds true
            while (rowIterator.hasNext()) {

                Row row = rowIterator.next();
                // For each row, iterate through all the
                // columns
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();

                    // Checking the cell type and format
                    // accordingly
                    switch (cell.getCellTypeEnum()) {
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "\t");
                            break;
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                        case BLANK:
                            System.out.print(" " + "\t");
                            break;
                        case _NONE:
                            System.out.print("None" + "\t");
                            break;
                        case BOOLEAN:
                            System.out.print("Bool" + "\t");
                            break;
                        case ERROR:
                            System.out.print("Error" + "\t");
                            break;
                        case FORMULA:
                            System.out.print(cell.getStringCellValue() + "\t");
                            break;
                    }
                }
                System.out.println("");
            }
            file.close();
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
// Main class
public class App {
    // Main driver method
    public static void main(String[] args) {
        excel obj = new excel();
        obj.getData();
        obj.getAverage();
        obj.writedata();
    }
}
