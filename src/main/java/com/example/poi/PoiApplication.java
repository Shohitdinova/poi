package com.example.poi;

import com.example.poi.exsel.read.Exsel_read;
import com.example.poi.exsel.write.Exsel;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;

import java.io.IOException;

@SpringBootApplication
public class PoiApplication {
    public static void main(String[] args) throws IOException {

        Exsel_read.readDataFromExcelFile("example.xlsx");





        // Word faylidan matn oqish
//        Word.readTextFromWordFile("example.docx");


        // word ga yozish
//        Scanner scanner = new Scanner(System.in);
//        System.out.println(" matnni kiriting :");
//        String s = scanner.nextLine();
//        WordService.createWordDocumentAndSaveToFile(Collections.singletonList(s));


       //  excel ga yozish
//        ConfigurableApplicationContext context = SpringApplication.run(PoiApplication.class, args);
//        Exsel exsel = context.getBean(Exsel.class);
//        exsel.excel();




















    }
}
