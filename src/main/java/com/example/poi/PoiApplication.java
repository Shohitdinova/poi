package com.example.poi;

import com.example.poi.exsel.Exsel;
import com.example.poi.word.WordService;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Scanner;

@SpringBootApplication
public class PoiApplication {
    public static void main(String[] args) {

























        // word ga yozish
//        Scanner scanner = new Scanner(System.in);
//        System.out.println(" matnni kiriting :");
//        String s = scanner.nextLine();
//        WordService.createWordDocumentAndSaveToFile(Collections.singletonList(s));


        // excel ga yozish
//        ConfigurableApplicationContext context = SpringApplication.run(PoiApplication.class, args);
//        Exsel exsel = context.getBean(Exsel.class);
//        exsel.excel();




















    }
}
