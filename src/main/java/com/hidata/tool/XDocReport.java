package com.hidata.tool;

import com.lowagie.text.Font;
import com.lowagie.text.pdf.BaseFont;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.xdocreport.itext.extension.font.IFontProvider;
import org.apache.commons.io.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.awt.*;
import java.io.*;

public class XDocReport {

    public static void main(String[] args) {
        String outPath = "E:\\workspace\\project\\playground\\wordTemplate\\doc\\ZhongYinOutFile.docx";
        wordConverterToPdf(outPath, outPath.replace("docx", "pdf"), "C:/Windows/Fonts/simkai.ttf", BaseFont.IDENTITY_H);
    }

    /**
     * 将word文档， 转换成pdf
     *
     * @param fontParam1 字体的路径
     * @param fontParam2 和fontParam2对应，fontParam1为路径时，fontParam2=BaseFont.IDENTITY_H，为itextasian-1.5.2.jar提供的字体时，fontParam2="UniGB-UCS2-H"
     * @param tmp        源为word文档， 必须为docx文档
     * @param target     目标输出
     * @throws Exception
     */
    public static void wordConverterToPdf(String tmp, String target, final String fontParam1, final String fontParam2) {
        InputStream sourceStream = null;
        OutputStream targetStream = null;
        XWPFDocument doc = null;
        try {
            sourceStream = new FileInputStream(tmp);
            targetStream = new FileOutputStream(target);
            doc = new XWPFDocument(sourceStream);
            PdfOptions options = PdfOptions.create();
            //中文字体处理
            options.fontProvider(new IFontProvider() {
                public Font getFont(String familyName, String encoding, float size, int style, Color color) {
                    try {
                        BaseFont bfChinese = BaseFont.createFont(fontParam1, fontParam2, BaseFont.NOT_EMBEDDED);
                        Font fontChinese = new Font(bfChinese, size, style, color);
                        if (familyName != null)
                            fontChinese.setFamily(familyName);
                        return fontChinese;
                    } catch (Exception e) {
                        e.printStackTrace();
                        return null;
                    }
                }
            });
            PdfConverter.getInstance().convert(doc, targetStream, options);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            IOUtils.closeQuietly(doc);
            IOUtils.closeQuietly(targetStream);
            IOUtils.closeQuietly(sourceStream);
        }
    }

}
