package com.example.poi.word.write;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

@Service
public class Word {

    public static void createWordDocumentAndSaveToFile(List<String> texts) {
        try (XWPFDocument document = new XWPFDocument();
             FileOutputStream out = new FileOutputStream("example.docx")) {

            System.out.println("saqlandi  !!!");

            for (String text : texts) {
                XWPFParagraph paragraph = document.createParagraph();
                XWPFRun run = paragraph.createRun();
                run.setText(text);
            }

            document.write(out);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
