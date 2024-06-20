package com.example.poi;

import com.example.poi.exsel.read.Exsel_read;
import com.example.poi.exsel.write.Exsel;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.VelocityEngine;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.ConfigurableApplicationContext;

import java.io.IOException;
import java.io.StringWriter;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Properties;

@SpringBootApplication
public class PoiApplication {

    public static void main(String[] args) throws IOException {



        String filePath = Paths.get("src", "main", "resources", "example.docx").toAbsolutePath().toString();
     //   String templatePath = Paths.get("src", "main", "resources", "read.vm").toAbsolutePath().toString();

        try {
            String content = DocsReader.readDocsFile(filePath);

            Properties props = new Properties();
            props.setProperty("resource.loader.file.class", "org.apache.velocity.runtime.resource.loader.FileResourceLoader");
            props.setProperty("resource.loader.file.path", Paths.get("src", "main", "resources").toAbsolutePath().toString());

            VelocityEngine ve = new VelocityEngine(props);
            ve.init();


            VelocityContext context = new VelocityContext();
            context.put("content", content);

            StringWriter writer = new StringWriter();
            ve.getTemplate("read.vm").merge(context, writer);

            System.out.println(writer);
        } catch (IOException e) {
            e.printStackTrace();
        }





        // exseldan oqish
 //       Exsel_read.readDataFromExcelFile("example.xlsx");

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
