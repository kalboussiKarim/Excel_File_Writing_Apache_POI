import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class CreateAndWriteToExcelFileWithIterator {

    public static void main(String[] args) {
        String[][] subscribersData = {
                {"Name", "Email", "Age"},
                {"David Johnson", "david.johnson@example.com", "28"},
                {"Emily Smith", "emily.smith@example.com", "34"},
                {"Michael Scott", "michael.scott@example.com", "45"},
                {"Anabel Lee", "anabel.lee@example.com", "29"}
        };

        String filePath = "./data/subscribers.xlsx";
        String sheetName = "subscribers 0";
        writeExcelFile(subscribersData, sheetName, filePath);
    }

    private static void writeExcelFile(String[][] data, String sheetName, String filePath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);

            int rowIndex = 0;
            for (String[] rowData : data) {
                Row row = sheet.createRow(rowIndex++);
                int colIndex = 0;
                for (String value : rowData) {
                    Cell cell = row.createCell(colIndex++);
                    setCellValue(cell, value);
                }
            }

            File outputFile = new File(filePath);
            outputFile.getParentFile().mkdirs();

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
                System.out.println("Excel file created successfully at: " + filePath);
            }

        } catch (IOException e) {
            System.err.println("Error writing Excel file: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static void setCellValue(Cell cell, String value) {
        try {
            cell.setCellValue(Integer.parseInt(value));
        } catch (NumberFormatException e) {
            cell.setCellValue(value);
        }
    }
}