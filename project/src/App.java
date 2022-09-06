import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

class excel {
    public void writedata() {
        try {
            System.out.println("WriteData() function exceuting");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void getData() {
        // Try block to check for exceptions
        try {
            // Reading file from local directory
            FileInputStream file = new FileInputStream(new File("data.xlsx"));

            // Create Workbook instance holding reference to
            // .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(1);

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
                            System.out.print("Blank" + "\t");
                            break;
                        case _NONE:
                            System.out.print("None" + "\t");
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
    }
}
