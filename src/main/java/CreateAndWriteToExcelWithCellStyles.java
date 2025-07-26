import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class CreateAndWriteToExcelWithCellStyles {

    public static void main(String[] args) {
        String[][] employeesData = {
                {"Name", "Email", "Age"},
                {"David Johnson", "david.johnson@example.com", "28"},
                {"Emily Smith", "emily.smith@example.com", "34"},
                {"Michael Scott", "michael.scott@example.com", "45"},
                {"Anabel Lee", "anabel.lee@example.com", "29"}
        };

        String filePath = "./data/employeesWithStyling.xlsx";
        String sheetName = "employeesStyling";
        writeExcelFile(employeesData, sheetName, filePath);
    }

    private static void writeExcelFile(String[][] data, String sheetName, String filePath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);

            // Create header style
            CellStyle headerStyle = workbook.createCellStyle();
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.CORAL.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            for (int i = 0; i < data.length; i++) {
                Row row = sheet.createRow(i);
                for (int j = 0; j < data[i].length; j++) {
                    Cell cell = row.createCell(j);
                    setCellValue(cell, data[i][j]);

                    // Apply header style to the first row only !!
                    if (i == 0) {
                        cell.setCellStyle(headerStyle);
                    }
                }
            }

            // Auto-size columns
            for (int col = 0; col < data[0].length; col++) {
                sheet.autoSizeColumn(col);
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
