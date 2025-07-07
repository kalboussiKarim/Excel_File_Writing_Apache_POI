import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;

import java.io.File;

public class CreateAndWriteToExcelFile {

    public static void main(String[] args) {
        String[][] employeesData = {
                {"Name", "Email", "Age"},
                {"David Johnson", "david.johnson@example.com", "28"},
                {"Emily Smith", "emily.smith@example.com", "34"},
                {"Michael Scott", "michael.scott@example.com", "45"},
                {"Anabel Lee", "anabel.lee@example.com", "29"}
        };

        String filePath = "./data/gymSubs.xlsx";
        String fileName = "Gym subscribers 0";
        writeExcelFile(employeesData, fileName, filePath);
    }

    private static void writeExcelFile(String[][] data, String sheetName, String filePath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);

            for (int i = 0; i < data.length; i++) {
                Row row = sheet.createRow(i);
                for (int j = 0; j < data[i].length; j++) {
                    Cell cell = row.createCell(j);
                    setCellValue(cell, data[i][j]);
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
