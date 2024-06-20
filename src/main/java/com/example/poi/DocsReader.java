package com.example.poi;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

public class DocsReader {
    public static String readDocsFile(String filePath) throws IOException {
        return new String(Files.readAllBytes(Paths.get(filePath)));
    }
}
