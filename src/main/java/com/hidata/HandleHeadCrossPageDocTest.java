package com.hidata;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

/**
 * @Author huabin
 * @DateTime 2022-07-28 15:08
 * @Desc 处理合并单元格表头跨页
 *
 */
public class HandleHeadCrossPageDocTest {

    public static void main(String[] args) throws IOException {
        File file = new File("/Users/huabin/workspace/playground/wordTemplate/doc/headCrossPageDoc.docx");//改成你本地文件所在目录
        FileInputStream fileInputStream = new FileInputStream(file);

        XWPFDocument document = new XWPFDocument(fileInputStream);
        List<IBodyElement> bodyElements = document.getBodyElements();  // 所有对象（段落+表格）

        int curT = 0;// 当前操作表格对象的索引
        int curP = 0;// 当前操作段落对象的索引
        for (int i = 0; i < bodyElements.size(); i++) {
            IBodyElement bodyElement = bodyElements.get(i);
            if (BodyElementType.TABLE.equals(bodyElement.getElementType())) {
                // 处理表格
                List<XWPFTable> tables = bodyElement.getBody().getTables();
                // 注意这里只能这么拿table，每次处理一个，因为bodyElements.size()肯定大于表数，在下面判空即可
                XWPFTable table = tables.get(curT);
                if (table != null) {
                    XWPFTableRow row = table.getRow(0);
                    List<XWPFTableCell> tableCells = row.getTableCells();
                    for (int i1 = 0; i1 < tableCells.size(); i1++) {
                        CTTcBorders tcBorders = tableCells.get(i1).getCTTc().getTcPr().getTcBorders();
                        System.out.println(tcBorders);
                    }
                    curT++;
                }
            }
        }
    }



}
