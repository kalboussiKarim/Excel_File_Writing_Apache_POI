import org.apache.poi.ss.usermodel.*;
        import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class CreateAndWriteToExcelFileWithCondCellStyles {

    public static void main(String[] args) {
        String[][] usersData = {
                {"Name", "Email", "Age"},
                {"Alice White", "alice.white@example.com", "17"},
                {"Bob Brown", "bob.brown@example.com", "22"},
                {"Charlie Gray", "charlie.gray@example.com", "31"},
                {"Diana Blue", "diana.blue@example.com", "45"},
                {"Ethan Black", "ethan.black@example.com", "19"},
                {"Fiona Green", "fiona.green@example.com", "28"},
                {"George Red", "george.red@example.com", "33"},
                {"Hannah Silver", "hannah.silver@example.com", "61"},
                {"Ian Gold", "ian.gold@example.com", "15"},
                {"Jenny Rose", "jenny.rose@example.com", "65"}
        };

        String filePath = "./data/users_conditional_styling.xlsx";
        String sheetName = "Users";

        writeExcelFile(usersData, sheetName, filePath);
    }

    private static void writeExcelFile(String[][] data, String sheetName, String filePath) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);

            // Header font & style
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFont(headerFont);
            headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            for (int i = 0; i < data.length; i++) {
                Row row = sheet.createRow(i);

                for (int j = 0; j < data[i].length; j++) {
                    Cell cell = row.createCell(j);
                    if (i != 0 && j == 2) { // Age column, not the header row
                        try {
                            cell.setCellValue(Integer.parseInt(data[i][j]));
                        } catch (NumberFormatException e) {
                            cell.setCellValue(data[i][j]); // fallback
                        }
                    } else {
                        cell.setCellValue(data[i][j]);
                    }

                    if (i == 0) {
                        cell.setCellStyle(headerStyle);
                    }
                }

                // Add "Category" column (skip header row)
                if (i == 0) {
                    Cell categoryHeader = row.createCell(data[0].length);
                    categoryHeader.setCellValue("Category");
                    categoryHeader.setCellStyle(headerStyle);
                } else {
                    int age = Integer.parseInt(data[i][2]);
                    String category = getCategory(age);

                    Cell categoryCell = row.createCell(data[0].length);
                    categoryCell.setCellValue(category);

                    // Apply background color based on category
                    CellStyle categoryStyle = workbook.createCellStyle();
                    categoryStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    categoryStyle.setFillForegroundColor(getCategoryColor(category));
                    categoryCell.setCellStyle(categoryStyle);
                }
            }

            // Auto-size all columns
            for (int i = 0; i <= data[0].length; i++) {
                sheet.autoSizeColumn(i);
            }

            File file = new File(filePath);
            file.getParentFile().mkdirs();

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
                System.out.println("Excel file created successfully at: " + filePath);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getCategory(int age) {
        if (age >= 13 && age <= 19) return "Teen";
        else if (age >= 20 && age <= 29) return "Young Adult";
        else if (age >= 30 && age <= 59) return "Adult";
        else return "Senior";
    }

    private static short getCategoryColor(String category) {
        return switch (category) {
            case "Teen" -> IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex();
            case "Young Adult" -> IndexedColors.LIGHT_GREEN.getIndex();
            case "Adult" -> IndexedColors.LIGHT_ORANGE.getIndex();
            case "Senior" -> IndexedColors.GREY_40_PERCENT.getIndex();
            default -> IndexedColors.WHITE.getIndex();
        };
    }
}
