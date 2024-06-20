package com.example.poi.exsel.write;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

@Service
public class Exsel {
     public void excel(){
         Scanner scanner = new Scanner(System.in);

         System.out.print("Nechta matn kiritmoqchisiz? ");
         int numberOfEntries = scanner.nextInt();
         scanner.nextLine();

         Workbook workbook = new XSSFWorkbook();
         Sheet sheet = workbook.createSheet("Example Sheet");

         for (int i = 0; i < numberOfEntries; i++) {
             System.out.print("Iltimos, " + (i + 1) + "-matnni kiriting: ");
             String inputText = scanner.nextLine();
             Row row = sheet.createRow(i+1);    //boyiga
             Cell cell = row.createCell(2);      // eniga
             cell.setCellValue(inputText);
         }

         try (FileOutputStream fileOut = new FileOutputStream("example.xlsx")) { // saqlanvotga fayl nomi
             workbook.write(fileOut);
         } catch (IOException e) {
             e.printStackTrace();
         }

         try {
             workbook.close();
         } catch (IOException e) {
             e.printStackTrace();
         }
         scanner.close();
         System.out.println("Matnlar muvaffaqiyatli yozildi va fayl saqlandi.");

     }


}
