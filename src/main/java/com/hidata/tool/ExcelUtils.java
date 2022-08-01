package com.hidata.tool;

import com.aspose.cells.Workbook;
import com.aspose.words.Document;
import com.aspose.words.License;
import com.aspose.words.SaveFormat;

import java.io.*;

/**
 * @Author huabin
 * @DateTime 2022-07-08 15:01
 * @Desc
 */
public class ExcelUtils {

    /**
     * 获取license
     *
     * @return
     */
    private static boolean getLicense() {
        boolean result = false;
        try {
            // 凭证
            String licenseStr =
                    "<License>\n" +
                            "  <Data>\n" +
                            "    <Products>\n" +
                            "      <Product>Aspose.Total for Java</Product>\n" +
                            "      <Product>Aspose.Cells for Java</Product>\n" +
                            "    </Products>\n" +
                            "    <EditionType>Enterprise</EditionType>\n" +
                            "    <SubscriptionExpiry>20991231</SubscriptionExpiry>\n" +
                            "    <LicenseExpiry>20991231</LicenseExpiry>\n" +
                            "    <SerialNumber>8bfe198c-7f0c-4ef8-8ff0-acc3237bf0d7</SerialNumber>\n" +
                            "  </Data>\n" +
                            "  <Signature>sNLLKGMUdF0r8O1kKilWAGdgfs2BvJb/2Xp8p5iuDVfZXmhppo+d0Ran1P9TKdjV4ABwAgKXxJ3jcQTqE/2IRfqwnPf8itN8aFZlV3TJPYeD3yWE7IT55Gz6EijUpC7aKeoohTb4w2fpox58wWoF3SNp6sK6jDfiAUGEHYJ9pjU=</Signature>\n" +
                            "</License>";
            InputStream license = new ByteArrayInputStream(licenseStr.getBytes("UTF-8"));
            License asposeLic = new License();
            asposeLic.setLicense(license);
            result = true;
        } catch (Exception e) {
            System.out.println("error:"+e);
        }
        return result;
    }

    /**
     * Word 2 pdf.
     *
     * @param resultFilePath   the pdf file path
     */
    public static void excelOpt(String xlsxFilePath, String resultFilePath) {
        FileOutputStream fileOS = null;
        // 验证License
//        if (!getLicense()) {
//            return;
//        }
        try {
            File file = new File(xlsxFilePath);
            // 读取word模板
            FileInputStream fileInputStream = new FileInputStream(file);
            Workbook workbook = new Workbook(xlsxFilePath);
            Document doc = new Document(fileInputStream);
            fileOS = new FileOutputStream(new File(resultFilePath));
            // 保存转换的pdf文件
            doc.save(fileOS, SaveFormat.PDF);
        } catch (Exception e) {
            System.out.println("error:"+e);
        } finally {
            try {
                if(fileOS != null){
                    fileOS.close();
                }
            } catch (IOException e) {
                System.out.println("error:"+e);
            }
        }
    }

    public static void main(String[] args) {
        excelOpt("/Users/huabin/workspace/playground/wordTemplate/doc/template.xlsx",
                "/Users/huabin/workspace/playground/wordTemplate/doc/result.xlsx");
    }

}
