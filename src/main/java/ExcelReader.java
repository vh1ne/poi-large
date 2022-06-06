import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.time.Duration;
import java.time.Instant;
import java.util.Iterator;
import java.util.Objects;

/**
 * vh
 */

public class ExcelReader {
    public static final String SAMPLE_XLS_FILE_PATH = "./sample-xls-file.xls";
    public static final String SAMPLE_XLSX_FILE_PATH = "E:\\Downloads\\Chrome Downloads\\Sample_900K.xlsx";
    //""E:\\Downloads\\Chrome Downloads\\Sample10K.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        Instant start = Instant.now();
        try
        {



            // Creating a Workbook from an Excel file (.xls or .xlsx)
           // XSSFWorkbook x= new XSSFWorkbook(new File(SAMPLE_XLSX_FILE_PATH));
            Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

           // SXSSFWorkbook workbook  = new SXSSFWorkbook(x,SXSSFWorkbook.DEFAULT_WINDOW_SIZE/* 100 */);
            // Retrieving the number of sheets in the Workbook
        //    System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

        // 1. You can obtain a sheetIterator and iterate over it
//        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
//        System.out.println("Retrieving Sheets using Iterator");
//     while (sheetIterator.hasNext()) {
//            Sheet sheet = sheetIterator.next();
//            System.out.println("=> " + sheet.getSheetName());
//        }
/*
        // 2. Or you can use a for-each loop
        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }

            // 3. Or you can use a Java 8 forEach wih lambda
            System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
            workbook.forEach(sheet -> {
                System.out.println("=> " + sheet.getSheetName());
            });


           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

            // Getting the Sheet at index zero
            Sheet sheet = workbook.getSheetAt(0);
           // Sheet sheet = x.getSheetAt(0);
            // Create a DataFormatter to format and get each cell's value as String
            DataFormatter dataFormatter = new DataFormatter();

       // 1. You can obtain a rowIterator and columnIterator and iterate over them
//        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
//        Iterator<Row> rowIterator = sheet.rowIterator();
//        while (rowIterator.hasNext()) {
//            Row row = rowIterator.next();
//
//            // Now let's iterate over the columns of the current row
//            Iterator<Cell> cellIterator = row.cellIterator();
//
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//                String cellValue = dataFormatter.formatCellValue(cell);
//                System.out.print(cellValue + "\t");
//            }
//            System.out.println();
//        }

        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }
            // 3. Or you can use Java 8 forEach loop with lambda
//            System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
//            sheet.forEach(row -> {
//                row.forEach(cell -> {
//                    printCellValue(cell);
//                });
//                System.out.println();
//            });

            // Closing the workbook
            workbook.close();
      //  x.close();
        }
        catch (Exception e)
        {
            System.out.println(e.getMessage());
        }
        Instant end = Instant.now();
        System.out.println(printExecutionTime(start,end));
    }


    public static String printExecutionTime(Instant start, Instant end)
    {
        return "Program  executed  in "+  (float) Duration.between(start, end).toMillis() / 1000  + " seconds." ;
    }

    private static void printCellValue(Cell cell) {
        switch (cell.getCellType()) {
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
}
