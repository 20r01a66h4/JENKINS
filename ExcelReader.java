package exautomation;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {
    private Workbook workbook;

    // Constructor to initialize the Excel file
    public ExcelReader(String filePath) throws IOException {
        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found at: " + filePath);
        }
        FileInputStream fileInputStream = new FileInputStream(file);
        workbook = new XSSFWorkbook(fileInputStream);
    }

    // Get data from a specific cell
    public String getCellData(int sheetIndex, int rowNumber, int columnNumber) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        if (sheet == null) {
            throw new RuntimeException("Sheet at index " + sheetIndex + " does not exist.");
        }

        Row row = sheet.getRow(rowNumber);
        if (row == null) {
            throw new RuntimeException("Row " + rowNumber + " does not exist in sheet " + sheetIndex);
        }

        Cell cell = row.getCell(columnNumber);
        if (cell == null) {
            throw new RuntimeException("Cell " + columnNumber + " does not exist in row " + rowNumber);
        }

        if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        }

        return cell.getStringCellValue();
    }

    // Close the workbook
    public void close() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }
}
