package maven.dustan.io.java_read_write_excel;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;

import org.apache.commons.net.telnet.TelnetClient;
import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * Hello world!
 *
 */
public class App 
{
	public static final String EXCEL_FILE_PATH = "./readable_writable_excel.xlsx";
    private static final Logger logger = LogManager.getLogger(App.class);
    public static void main( String[] args ) throws Exception {
        PropertyConfigurator.configure("log4j.properties");
        // System.out.println( "Hello World!" );
    	// Creating a Workbook from an Excel file (.xls or .xlsx)
//        BasicConfigurator.configure();

        Workbook workbook = WorkbookFactory.create(new File(EXCEL_FILE_PATH));


        // 1. You can obtain a sheetIterator and iterate over it

        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
//        System.out.println("Retrieving Sheets using Iterator");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
//            System.out.println("=> " + sheet.getSheetName());
        }

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // 1. You can obtain a rowIterator and columnIterator and iterate over them
//        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        int index_row = 0;
        logger.info("START");
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();
            int index_col = 0;
            String ip = "";
            int port = 0;
            while (index_col < 2) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
//                System.out.print(cellValue + "\t");
                index_col ++;
                if (cellValue == ""){
                    break;
                }

                if(index_col == 1){
                    ip = cellValue;
                }
                if(index_col == 2){
                    port = Integer.parseInt(cellValue);
                }

            }
            System.out.println(telnet(ip,port) + "  ip "+ip+" port "+port);
            logger.info(telnet(ip,port) + "  ip "+ip+" port "+port);
            index_row++;
            if (index_row == 2){
                break;
            }
        }
        logger.info("END");

        // Closing the workbook
        workbook.close();
    }
    
    private static void printCellValue(Cell cell) throws IOException {

        switch (cell.getCellTypeEnum()) {
            case BOOLEAN:
                System.out.print(cell.getBooleanCellValue());
                break;
            case STRING:
                System.out.print(cell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    System.out.print(cell.getDateCellValue());
                } else {
                    System.out.print(cell.getNumericCellValue());
                }
                break;
            case FORMULA:
                System.out.print(cell.getCellFormula());
                break;
            case BLANK:
                System.out.print("");
                break;
            default:
                System.out.print("");
        }


        System.out.print("\t");
    }
    public  static boolean telnet(String ip, int port) throws Exception {
        TelnetClient telnetClient = new TelnetClient("vt200");
        telnetClient.setDefaultTimeout(5000);
        try {
            telnetClient.connect(ip,port);
        } catch (Exception e) {
            return false;
        }
        return true;
    }
}
