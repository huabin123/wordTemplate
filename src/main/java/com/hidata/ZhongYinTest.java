package com.hidata;

import com.hidata.tool.LibreOffice;
import com.hidata.tool.WordUtils;
import com.hidata.tool.XDocReport;
import com.lowagie.text.pdf.BaseFont;
import org.apache.poi.util.IOUtils;
import org.apache.xmlbeans.XmlException;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.HashMap;

public class ZhongYinTest {

    public static void main(String[] args) throws IOException, XmlException {
//        HashMap<String, String> map = new HashMap<>();
//        File file = new File("E:\\workspace\\project\\playground\\wordTemplate\\doc\\附件七：定期报告需求模板(6).docx");//改成你本地文件所在目录
////        File file = new File("C:\\\\Users\\\\me\\\\Desktop\\\\附件七：定期报告需求模板(6).docx");//改成你本地文件所在目录
//
//        // 读取word模板
//        FileInputStream fileInputStream = new FileInputStream(file);
//        ByteArrayInputStream arrayInputStream = new WordUtils().replaceDocument(map, fileInputStream);
//
//        //生成文件
//        File outputFile=new File("E:\\workspace\\project\\playground\\wordTemplate\\doc\\ZhongYinOutFile.docx");//改成你本地文件所在目录
//        Files.copy(arrayInputStream, outputFile.toPath(), StandardCopyOption.REPLACE_EXISTING);
//        IOUtils.closeQuietly(arrayInputStream);

        // 使用xdocreport word转pdf
        String outPath = "E:\\workspace\\project\\playground\\wordTemplate\\doc\\ZhongYinOutFile.docx";
        new XDocReport().wordConverterToPdf(outPath, outPath.replace("docx", "pdf"), "C:/Windows/Fonts/simkai.ttf", BaseFont.IDENTITY_H);

//        // 使用libreOffice word转pdf
//        String outPath = "E:\\workspace\\project\\playground\\wordTemplate\\doc\\ZhongYinOutFile.docx";
//        new LibreOffice().wordConverterToPdf(outPath);

    }
}

