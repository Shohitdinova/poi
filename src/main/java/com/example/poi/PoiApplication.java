package com.example.poi;

import com.example.poi.exsel.read.Exsel_read;
import com.example.poi.exsel.write.Exsel;
import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.VelocityEngine;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;

import java.io.IOException;
import java.io.StringWriter;
import java.util.List;
import java.util.Properties;

@SpringBootApplication
public class PoiApplication {

    public static void main(String[] args) throws IOException {


                                                          //.vm bilan excelga  yozsih
        Exsel exsel= new Exsel();
        exsel.excel();

                                                           //.vm bilan excel oqish
//        Properties props = new Properties();
//        props.setProperty("resource.loader", "classpath");
//        props.setProperty("classpath.resource.loader.class", "org.apache.velocity.runtime.resource.loader.ClasspathResourceLoader");
//        VelocityEngine velocityEngine = new VelocityEngine(props);
//        Template template = velocityEngine.getTemplate("templates/template.vm");
//        String excelFilePath = "src/main/resources/excel/example.xlsx";
//        List<List<String>> data = ExcelReader.readExcel(excelFilePath);
//        VelocityContext context = new VelocityContext();
//        context.put("data", data);
//        StringWriter writer = new StringWriter();
//        template.merge(context, writer);
//        System.out.println(writer);

                                                           // exseldan oqish
 //       Exsel_read.readDataFromExcelFile("example.xlsx");

                                                           // Word faylidan matn oqish
//        Word.readTextFromWordFile("example.docx");

                                                             //  excel ga yozish
//        ConfigurableApplicationContext context = SpringApplication.run(PoiApplication.class, args);
//        Exsel exsel = context.getBean(Exsel.class);
//        exsel.excel();

                                                             // word ga yozish
//        Scanner scanner = new Scanner(System.in);
//        System.out.println(" matnni kiriting :");
//        String s = scanner.nextLine();
//        WordService.createWordDocumentAndSaveToFile(Collections.singletonList(s));




    }
}
