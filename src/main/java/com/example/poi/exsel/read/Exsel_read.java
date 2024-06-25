package com.example.poi.exsel.read;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class Exsel_read {
    public static void readDataFromExcelFile(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream((filePath));
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            System.out.println("======== Ma'lumotlar =======");
            for (Row row : sheet) {
                for (Cell cell : row) {
                    switch (cell.getCellType()) {
                        case STRING:
                            System.out.print(cell.getStringCellValue() + "  \t ");
                            break;
                        case NUMERIC:
                            System.out.print(cell.getNumericCellValue() + "  \t ");
                            break;
                        case BOOLEAN:
                            System.out.print(cell.getBooleanCellValue() + "   \t ");
                            break;
                        case BLANK:
                            System.out.print("<BLANK>\t");
                            break;
                        default:
                            System.out.print("<UNKNOWN>\t");
                    }
                }
                System.out.println();
            }
        }
    }

    public static List<List<String>> readExcel(String filePath) {
        List<List<String>> data = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath)) {
            try (Workbook workbook = new XSSFWorkbook(fis)) {

                Sheet sheet = workbook.getSheetAt(0); // Birinchi varaqni o'qish
                for (Row row : sheet) {
                    List<String> rowData = new ArrayList<>();
                    for (Cell cell : row) {
                        switch (cell.getCellType()) {
                            case STRING:
                                rowData.add(cell.getStringCellValue());
                                break;
                            case NUMERIC:
                                rowData.add(String.valueOf(cell.getNumericCellValue()));
                                break;
                            default:
                                rowData.add("");
                        }
                    }
                    data.add(rowData);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return data;
    }

}
