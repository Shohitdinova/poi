package com.example.poi.exsel.write;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.VelocityEngine;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
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

         try (FileOutputStream fileOut = new FileOutputStream("src/main/resources/excel/example.xlsx")) { // saqlanayotgan fayl nomi
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







    public void excel_vm() {
        Scanner scanner = new Scanner(System.in);

        System.out.print("Nechta matn kiritmoqchisiz? ");
        int numberOfEntries = scanner.nextInt();
        scanner.nextLine();

        List<List<String>> data = new ArrayList<>();
        for (int i = 0; i < numberOfEntries; i++) {
            List<String> row = new ArrayList<>();
            System.out.print("Iltimos, " + (i + 1) + "-matnni kiriting: ");
            String inputText = scanner.nextLine();
            row.add(inputText);
            data.add(row);
        }
        scanner.close();

        Properties props = new Properties();
        props.setProperty("resource.loader", "classpath");
        props.setProperty("classpath.resource.loader.class", "org.apache.velocity.runtime.resource.loader.ClasspathResourceLoader");

        VelocityEngine velocityEngine = new VelocityEngine(props);

        Template template = velocityEngine.getTemplate("templates/template.vm");

        VelocityContext context = new VelocityContext();
        context.put("data", data);

        StringWriter writer = new StringWriter();
        template.merge(context, writer);

        System.out.println(writer);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Example Sheet");

        String[] lines = writer.toString().split("\n");
        for (int i = 0; i < lines.length; i++) {
            Row row = sheet.createRow(i);
            String[] cells = lines[i].split(" ");
            for (int j = 0; j < cells.length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(cells[j]);
            }
        }

        try (FileOutputStream fileOut = new FileOutputStream("src/main/resources/excel/example.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Matnlar muvaffaqiyatli yozildi va fayl saqlandi.");
    }
}