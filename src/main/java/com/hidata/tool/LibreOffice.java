package com.hidata.tool;

import java.io.File;
import java.io.IOException;

/**
 * 使用libreoffice来进行转pdf
 */
public class LibreOffice {

    public LibreOffice(){}

    public void wordConverterToPdf(String docxPath) throws IOException {
        File file = new File(docxPath);
        String path = file.getParent();
        try {
            String osName = System.getProperty("os.name");
            String command = "";
            if (osName.contains("Windows")) {
                command = "soffice --convert-to pdf  -outdir " + path + " " + docxPath;
            } else {
                command = "doc2pdf --output=" + path + File.separator + file.getName().replaceAll(".(?i)docx", ".pdf") + " " + docxPath;
            }
            String result = CommandExecute.executeCommand(command);
            if (result.equals("") || result.contains("writer_pdf_Export")) {
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        }
    }
}
