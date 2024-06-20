package com.example.poi.word.read;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.springframework.stereotype.Service;

import java.io.FileInputStream;
import java.io.IOException;

@Service
public class Word_read {
    public static void readTextFromWordFile(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            // Word faylidagi barcha matnlarni oqish va konsolga chiqarish
            System.out.println("============  Matn:  =================");
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                String text = paragraph.getText();
                if (!text.isEmpty()) {
                    System.out.println(text);
                }
            }
        }


    }}
