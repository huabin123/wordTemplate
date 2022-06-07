package com.hidata.tool;


import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WordUtils {

    private static XWPFDocument document;

//    public XWPFDocument getDocument() {
//        return document;
//    }
//
//    public void setDocument(XWPFDocument document) {
//        this.document = document;
//    }

    public WordUtils() {

    }

    public static List<Map<String, Object>> getIndexData(Map<String, String> parametersMap){
        HashMap<String, String> singleIndicatorMap = new HashMap<>();
        HashMap<String, Object> multiIndicatorMap = new HashMap<>();
        List resultList = new ArrayList<>();
        int tableIndex = 0;  // 要删除的配置表的下标

        List<IBodyElement> bodyElements = document.getBodyElements();  // 所有对象（段落+表格）
        for (int i = 0; i < bodyElements.size(); i++) {
            IBodyElement bodyElement = bodyElements.get(i);
            if (BodyElementType.TABLE.equals(bodyElement.getElementType())) {
                XWPFTable table = bodyElement.getBody().getTables().get(0);  // 第一张表就是配置表,第一列是key，第二列是value
                if (table != null) {
                    List<XWPFTableRow> tableRows = table.getRows();
                    for (XWPFTableRow row : tableRows) {
                        List<XWPFTableCell> tableCells = row.getTableCells();
                        String key = tableCells.get(0).getText();
                        String formula = tableCells.get(1).getText();
                        String singleIndicator = "";
                        List<List<String>> multi = new ArrayList<>();
                        // 根据公式请求指标数据
                        try{
                            if (formula.startsWith("=SF_VFUN")) {
                                singleIndicatorMap.put(key, singleIndicator);
                            }else if (formula.startsWith("=SF_MDFUN")){
                                multiIndicatorMap.put(key, multi);
                            }
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                    }
                    resultList.add(singleIndicatorMap);
                    resultList.add(multiIndicatorMap);
                    break;
                }
            }
            tableIndex++;
        }
        // 删除第一张配置表
        document.removeBodyElement(tableIndex);
        return resultList;
    }

    public static ByteArrayInputStream replaceDocument(Map<String, String> parametersMap,
                                                       InputStream inputStream) throws XmlException, IOException {

        document = new XWPFDocument(inputStream);

        // 读取第一张配置表，保存指标，读取完成后删除第一张表
        List<Map<String, Object>> dataList = getIndexData(parametersMap);
        Map<String, Object> singleIndicatorMap = dataList.get(0);
        singleIndicatorMap.put("D0000106001", "报告期内产品流动性平稳运行，规模保持稳定，产品管理人通过合理安排资产配置结构，保持一定比例的高流动性资产，控制资产久期、杠杆融资比例，管控产品流动性风险。");
        singleIndicatorMap.put("CASH_MGMT_CLS_FLAG", "1");
        Map<String, Object> multiIndicatorMap = dataList.get(1);

        List<List> testList = new ArrayList<>();
        // 造数据
        List<String> row0 = new ArrayList<>();
        row0.add("阿萨德\r\n九分裤垃圾收代理费空间啊");
        row0.add("阿萨德和房价开始大家付款哈实践课程vhajksdhfjkasdhfk");
        row0.add("阿萨德和房价开始大家付款哈实践课程vhajksdqwdasdasd奥术大师大所多阿是大师大师的奥术大师大所大所大所多hfjkasdh萨德和房价开始大家付款哈实践课程vhajksdqwdasdasd奥术大师大所多阿是大师大师的奥术大师大所大所大所多hfjk萨德和房价开始大家付款哈实践课程vhajksdqwdasdasd奥术大师大所多阿是大师大师的奥术大师大所大所大所多hfjk萨德和房价开始大家付款哈实践课程vhajksdqwdasdasd奥术大师大所多阿是大师大师的奥术大师大所大所大所多hfjk萨德和房价开始大家付款哈实践课程vhajksdqwdasdasd奥术大师大所多阿是大师大师的奥术大师大所大所大所多hfjkfk");
        row0.add("阿萨德和房价开始大家付款哈实践课程vhajksdhfjkasdhfk");
        row0.add("阿萨德和房价开始大家付款哈实践课程vhajksdhfjkasdhfk");

        List<String> row1 = new ArrayList<>();
        row1.add("阿萨德\n九分裤垃圾收代理费空间啊");
        row1.add("阿萨德和房价开始大家付款哈实践课程vhajksdhfjkasdhfk");
        row1.add("阿萨德和房价开始大家付款哈实践课程vhajksdhfjkasdhfk");
        row1.add("阿萨德和房价开始大家付款哈实践课程vhajksdhfjkasdhfk");
        row1.add("阿萨德和房价开始大家付款哈实践课程vhajksdhfjkasdhfk");

        List<String> row2 = new ArrayList<>();
        row2.add("阿萨德\n九分裤垃圾收代理费空间啊");
        row2.add("阿萨德和房价开始大家付款哈实践课程vhajksdhfjkasdhfk");
        row2.add("阿萨德和房价开始大家付款哈实践课程vhajksdhfjkasdhfk");
        row2.add("阿萨德和房价开始大家付款哈实践课程vhajksdhfjkasdhfk");
        row2.add("阿萨德和房价开始大家付款哈实践课程vhajksdhfjkasdhfk");

        testList.add(row0);
        testList.add(row1);
        testList.add(row2);

        multiIndicatorMap.put("D00001非标资产投资情况", testList);

        // 处理if endif条件，除去不需要的表格和模板
        removeDocumentByCondition(singleIndicatorMap);

        String regEx = "\\$\\{(.*?)\\}";
        Pattern pattern = Pattern.compile(regEx);
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
                    for (int k = 0; k < table.getRows().size(); k++) {
                        for (int l = 0; l < table.getRow(k).getTableCells().size(); l++) {
                            String cellString = table.getRow(k).getTableCells().get(l).getText();
                            if (cellString.contains("${")) {  // 说明需要替换内容
                                // 获取key
                                Matcher matcher = pattern.matcher(cellString);
                                if (matcher.find()) {
                                    String keyString = matcher.group(1);
                                    if (singleIndicatorMap.containsKey(keyString)) {  // 单维度指标
                                        table.getRow(k).getTableCells().get(l).removeParagraph(0);
                                        table.getRow(k).getTableCells().get(l).setText((String) singleIndicatorMap.get(keyString));
                                    } else if (multiIndicatorMap.containsKey(keyString)) {  // 多维指标
                                        List<List<String>> sourceDataList = (List<List<String>>) multiIndicatorMap.get(keyString);
                                        setSFMDFUNValue(table, k, l, sourceDataList);
                                    } else {
                                        // 无此指标
                                        table.getRow(k).getTableCells().get(l).removeParagraph(0);
                                        table.getRow(k).getTableCells().get(l).setText("");
                                    }
                                }
                            }
                        }
                    }
                    curT++;
                }
            }
            else if (BodyElementType.PARAGRAPH.equals(bodyElement.getElementType())) {// 处理段落
                XWPFParagraph ph = bodyElement.getBody().getParagraphArray(curP);
                if (ph != null) {
                    replaceParagraph(ph, singleIndicatorMap);
                    curP++;
                }
            }
        }

        // 生成input io流
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        document.write(baos);
        return new ByteArrayInputStream(baos.toByteArray());
    }

    private static void forverseTableCells() {
        for(XWPFTable table : document.getTables()) {//表格
            for(XWPFTableRow row : table.getRows()) {//行
                for(XWPFTableCell cell : row.getTableCells()) {//单元格 : 直接cell.setText()只会把文字加在原有的后面，删除不了文字
                    addBreakInCell(cell);
                }
            }
        }
    }

    private static void addBreakInCell(XWPFTableCell cell) {
        if(cell.getText() != null && cell.getText().contains("\n")) {

            for (XWPFParagraph p : cell.getParagraphs()) {
                for (XWPFRun run : p.getRuns()) {//XWPFRun对象定义具有一组公共属性的文本区域
                    if (run.getText(0) != null && run.getText(0).contains("\n")) {
                        String[] lines = run.getText(0).split("\n");
                        if (lines.length > 0) {
                            run.setText(lines[0], 0); // set first line into XWPFRun
                            for (int i = 1; i < lines.length; i++) {
                                // add break and insert new text
                                run.addCarriageReturn();//中断
//                                    run.addCarriageReturn();//回车符，但是不起作用
                                run.setText(lines[i]);
                            }
                        }
                    }
                }
            }
        }
    }

    /**
     * 填充多维表格数据
     * @param table 表格
     * @param rowStart 多行指标所在行
     * @param colStart 多维指标所在列
     * @param rowList 多维指标数据
     */
    private static void setSFMDFUNValue(XWPFTable table, int rowStart, int colStart, List<List<String>> rowList){
        if (rowList.size()==0) {
            // 多维指标为空，把这一行全部置空
            List<XWPFTableCell> tableCells = table.getRow(rowStart).getTableCells();
            for (int i = colStart; i < tableCells.size(); i++) {
                tableCells.get(i).removeParagraph(0);
                tableCells.get(i).setText("");
                // 设置居中
                CTTc cttc = tableCells.get(i).getCTTc();
                CTTcPr ctPr = cttc.addNewTcPr();
                ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
                cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
            }
        } else {
            CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
            CTBorder hBorder = borders.addNewInsideH();
            hBorder.setVal(STBorder.Enum.forString("single"));  // 线条类型
            hBorder.setSz(new BigInteger("1")); // 线条大小
            hBorder.setColor("000000"); // 设置颜色

            CTBorder vBorder = borders.addNewInsideV();
            vBorder.setVal(STBorder.Enum.forString("single"));
            vBorder.setSz(new BigInteger("1"));
            vBorder.setColor("000000");

            CTBorder lBorder = borders.addNewLeft();
            lBorder.setVal(STBorder.Enum.forString("single"));
            lBorder.setSz(new BigInteger("1"));
            lBorder.setColor("000000");

            CTBorder rBorder = borders.addNewRight();
            rBorder.setVal(STBorder.Enum.forString("single"));
            rBorder.setSz(new BigInteger("1"));
            rBorder.setColor("000000");

            CTBorder tBorder = borders.addNewTop();
            tBorder.setVal(STBorder.Enum.forString("single"));
            tBorder.setSz(new BigInteger("1"));
            tBorder.setColor("000000");

            CTBorder bBorder = borders.addNewBottom();
            bBorder.setVal(STBorder.Enum.forString("single"));
            bBorder.setSz(new BigInteger("1"));
            bBorder.setColor("000000");

            for (int i = 0; i < rowList.size(); i++) {
                XWPFTableRow tableRow=table.getRow(i+rowStart)==null?table.createRow():table.getRow(i+rowStart);
                for (int j = 0; j < rowList.get(i).size(); j++) {
                    XWPFTableCell tableCell=tableRow.getCell(j+colStart)==null?tableRow.createCell():tableRow.getCell(j+colStart);
                    tableCell.removeParagraph(0);
                    String rowText = rowList.get(i).get(j);

                    XWPFParagraph paragraph1 = tableCell.addParagraph();

                    if(rowText.contains("\n")) {
                        String[] text = rowText.split("\n");
                        paragraph1.insertNewRun(0).setText(text[0]);

                        int xx = 1;
                        for(int p=1;p<text.length;p++){
                            // add break and insert new text
                            paragraph1.insertNewRun(xx).addBreak();//中断
                            paragraph1.insertNewRun(xx+1).setText(text[p]);
                            xx = xx + 2;
                        }
                    }else {
                        tableCell.setText(rowText);
                    }

                    // 设置居中
                    CTTc cttc = tableCell.getCTTc();
                    CTTcPr ctPr = cttc.addNewTcPr();
                    ctPr.addNewVAlign().setVal(STVerticalJc.CENTER);
                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);

                }
            }
        }

    }

    public static void replaceParagraph(XWPFParagraph xWPFParagraph, Map<String, Object> parametersMap) {
        List<XWPFRun> runs = xWPFParagraph.getRuns();
        String xWPFParagraphText = xWPFParagraph.getText();
        String regEx = "\\$\\{.+?\\}";
        Pattern pattern = Pattern.compile(regEx);
        Matcher matcher = pattern.matcher(xWPFParagraphText);//正则匹配字符串{****}

        if (matcher.find()) {
            // 查找到有标签才执行替换
            int beginRunIndex = xWPFParagraph.searchText("${", new PositionInParagraph()).getBeginRun();// 标签开始run位置
            int endRunIndex = xWPFParagraph.searchText("}", new PositionInParagraph()).getEndRun();// 结束标签
            StringBuffer key = new StringBuffer();

            if (beginRunIndex == endRunIndex) {
                // {**}在一个run标签内
                XWPFRun beginRun = runs.get(beginRunIndex);
                String beginRunText = beginRun.text();

                int beginIndex = beginRunText.indexOf("${");
                int endIndex = beginRunText.indexOf("}");
                int length = beginRunText.length();

                if (beginIndex == 0 && endIndex == length - 1) {
                    // 该run标签只有{**}
                    XWPFRun insertNewRun = xWPFParagraph.insertNewRun(beginRunIndex);
                    insertNewRun.getCTR().setRPr(beginRun.getCTR().getRPr());
                    // 设置文本
                    key.append(beginRunText.substring(2, endIndex));
                    insertNewRun.setText(getValueBykey(key.toString(),parametersMap));
                    xWPFParagraph.removeRun(beginRunIndex + 1);
                } else {
                    // 该run标签为**{**}** 或者 **{**} 或者{**}**，替换key后，还需要加上原始key前后的文本
                    XWPFRun insertNewRun = xWPFParagraph.insertNewRun(beginRunIndex);
                    insertNewRun.getCTR().setRPr(beginRun.getCTR().getRPr());
                    // 设置文本
                    key.append(beginRunText.substring(beginRunText.indexOf("${")+2, beginRunText.indexOf("}")));
                    String textString=beginRunText.substring(0, beginIndex) + getValueBykey(key.toString(),parametersMap)
                            + beginRunText.substring(endIndex + 1);
                    insertNewRun.setText(textString);
                    xWPFParagraph.removeRun(beginRunIndex + 1);
                }

            }else {
                // {**}被分成多个run

                //先处理起始run标签,取得第一个{key}值
                XWPFRun beginRun = runs.get(beginRunIndex);
                String beginRunText = beginRun.text();
                int beginIndex = beginRunText.indexOf("${");
                if (beginRunText.length()>1  ) {
                    key.append(beginRunText.substring(beginIndex+1));
                }
                ArrayList<Integer> removeRunList = new ArrayList<>();//需要移除的run
                //处理中间的run
                for (int i = beginRunIndex + 1; i < endRunIndex; i++) {
                    XWPFRun run = runs.get(i);
                    String runText = run.text();
                    key.append(runText);
                    removeRunList.add(i);
                }

                // 获取endRun中的key值
                XWPFRun endRun = runs.get(endRunIndex);
                String endRunText = endRun.text();
                int endIndex = endRunText.indexOf("}");
                //run中**}或者**}**
                if (endRunText.length()>1 && endIndex!=0) {
                    key.append(endRunText.substring(0,endIndex));
                }
                //先处理开始标签
                if (beginRunText.length()==2 ) {
                    // run标签内文本{
                    XWPFRun insertNewRun = xWPFParagraph.insertNewRun(beginRunIndex);
                    insertNewRun.getCTR().setRPr(beginRun.getCTR().getRPr());
                    // 设置文本
                    insertNewRun.setText(getValueBykey(key.toString(),parametersMap));
                    xWPFParagraph.removeRun(beginRunIndex + 1);//移除原始的run
                }else {
                    // 该run标签为**{**或者 {** ，替换key后，还需要加上原始key前的文本
                    XWPFRun insertNewRun = xWPFParagraph.insertNewRun(beginRunIndex);
                    insertNewRun.getCTR().setRPr(beginRun.getCTR().getRPr());
                    // 设置文本
                    String textString=beginRunText.substring(0,beginRunText.indexOf("${")+1)+getValueBykey(key.toString(),parametersMap);
                    insertNewRun.setText(textString);
                    xWPFParagraph.removeRun(beginRunIndex + 1);//移除原始的run
                }

                //处理结束标签
                if (endRunText.length()==1 ) {
                    // run标签内文本只有}
                    XWPFRun insertNewRun = xWPFParagraph.insertNewRun(endRunIndex);
                    insertNewRun.getCTR().setRPr(endRun.getCTR().getRPr());
                    // 设置文本
                    insertNewRun.setText("");
                    xWPFParagraph.removeRun(endRunIndex + 1);//移除原始的run

                }else {
                    // 该run标签为**}**或者 }** 或者**}，替换key后，还需要加上原始key后的文本
                    XWPFRun insertNewRun = xWPFParagraph.insertNewRun(endRunIndex);
                    insertNewRun.getCTR().setRPr(endRun.getCTR().getRPr());
                    // 设置文本
                    String textString=endRunText.substring(endRunText.indexOf("}")+1);
                    insertNewRun.setText(textString);
                    xWPFParagraph.removeRun(endRunIndex + 1);//移除原始的run
                }

                //处理中间的run标签
                for (int i = 0; i < removeRunList.size(); i++) {
                    XWPFRun xWPFRun = runs.get(removeRunList.get(i));//原始run
                    XWPFRun insertNewRun = xWPFParagraph.insertNewRun(removeRunList.get(i));
                    insertNewRun.getCTR().setRPr(xWPFRun.getCTR().getRPr());
                    insertNewRun.setText("");
                    xWPFParagraph.removeRun(removeRunList.get(i) + 1);//移除原始的run
                }

            }// 处理${**}被分成多个run

            replaceParagraph( xWPFParagraph, parametersMap);

        }//if 有标签

    }

    private static String getValueBykey(String key, Map<String, Object> map) {
        String returnValue="";
        if (key != null) {
            try {
                returnValue=(String) map.get(key)!=null ? (String)map.get(key).toString() : "";
            } catch (Exception e) {
                // TODO: handle exception
                System.out.println("key:"+key+"***"+e);
                returnValue="";
            }

        }
        return returnValue;
    }

    private static boolean isIfExists(){
        List<IBodyElement> bodyElements = document.getBodyElements();// 所有对象（段落+表格）
        int templateBodySize = bodyElements.size();// 标记模板文件（段落+表格）总个数
        int curP = 0;// 当前操作段落对象的索引
        for (int a = 0; a < templateBodySize; a++) {
            IBodyElement body = bodyElements.get(a);
            if (BodyElementType.PARAGRAPH.equals(body.getElementType())) {
                XWPFParagraph ph = body.getBody().getParagraphArray(curP);
                if (ph != null) {
                    String text = ph.getText();
                    if (text.contains("IF")) {
                        return true;
                    }else {
                        curP++;
                    }
                }
            }
        }
        return false;
    }

    public static void removeDocumentByCondition(Map<String, Object> parametersMap){

        while (isIfExists()){
            removeDocument(parametersMap);
        }

    }

    // 循环调用直到文档中不包含含有if的段落
    public static void removeDocument(Map<String, Object> parametersMap) {

        List<IBodyElement> bodyElements = document.getBodyElements();// 所有对象（段落+表格）
        int templateBodySize = bodyElements.size();// 标记模板文件（段落+表格）总个数

        String regEx = "\\$(.*?)=\"(.*?)\"";
        Pattern pattern = Pattern.compile(regEx);
        int curP = 0;// 当前操作段落对象的索引
        int curT = 0;// 当前操作表格对象的索引
        for (int a = 0; a < templateBodySize; a++) {
            IBodyElement body = bodyElements.get(a);
            if (BodyElementType.TABLE.equals(body.getElementType())){
                // copy表格
                List<XWPFTable> tables = body.getBody().getTables();
                XWPFTable table = tables.get(curT);
                if (table != null) {
                    copyTableInDocFooter(table);
                    curT++;
                }
            } else if (BodyElementType.PARAGRAPH.equals(body.getElementType())) {
                XWPFParagraph ph = body.getBody().getParagraphArray(curP);
                if (ph != null) {
                    String text = ph.getText();
                    if (text.contains("IF")) {
                        // 需要考虑到嵌套if的情况，第一个if条件满足
                        Matcher matcher = pattern.matcher(text);
                        if (matcher.find()) {
                            String keyString = matcher.group(1);
                            String valueString = matcher.group(2);
                            String completeString = matcher.group(0);
                            if (valueString.equals(parametersMap.get(keyString))) {
                                // 符合条件，拷贝到endif之间的段落和表格
                                List<IBodyElement> subElements = new ArrayList<>(bodyElements.subList(a+1, bodyElements.size()));
                                for (IBodyElement subElement : subElements) {
                                    if (BodyElementType.TABLE.equals(subElement.getElementType())){
                                        XWPFTable table = body.getBody().getTableArray(curT+1);
                                        if (table != null) {
                                            copyTableInDocFooter(table);
                                            curT++;
                                            a++;
                                        }
                                    } else if (BodyElementType.PARAGRAPH.equals(subElement.getElementType())){
                                        XWPFParagraph innerPh = body.getBody().getParagraphArray(curP+1);
                                        if (innerPh != null) {
                                            String innerPhText = innerPh.getText();
                                            if (innerPhText.startsWith("=ENDIF("+completeString)){
                                                curP++;
                                                a++;
                                                break;
                                            } else {
                                                copyParagraphInDocFooter(innerPh);
                                            }
                                            curP++;
                                            a++;
                                        }
                                    }
                                }

                            } else {
                                // 不符合条件，if和endif区间的表格段落不需要复制
                                if (text.startsWith("=IF")) {
                                    List<IBodyElement> subElements = new ArrayList<>(bodyElements.subList(a+1, bodyElements.size()));
                                    for (IBodyElement subElement : subElements) {
                                        if (BodyElementType.TABLE.equals(subElement.getElementType())){
                                            XWPFTable table = body.getBody().getTableArray(curT+1);
                                            if (table != null) {
                                                curT++;
                                                a++;
                                            }
                                        } else if (BodyElementType.PARAGRAPH.equals(subElement.getElementType())){
                                            XWPFParagraph innerPh = body.getBody().getParagraphArray(curP+1);
                                            if (innerPh != null) {
                                                String innerPhText = innerPh.getText();
                                                if (innerPhText.startsWith("=ENDIF("+completeString)){
                                                    curP++;
                                                    a++;
                                                    break;
                                                }
                                                curP++;
                                                a++;
                                            }
                                        }
                                    }
                                }

                            }
                        }
                    }else{
                        // 由于if前都要空一行以正确识别if条件，所以如果下一个element是段落，且段落里含有if，那么本空行就不需要复制
                        boolean copyFlag = true;
                        IBodyElement iBodyElement = bodyElements.get(a + 1);
                        if (BodyElementType.PARAGRAPH.equals(iBodyElement.getElementType())) {
                            XWPFParagraph paragraph = iBodyElement.getBody().getParagraphArray(curP + 1);
                            if (paragraph.getText().contains("IF")) {
                                copyFlag = false;
                            }
                        }
                        if (copyFlag) {
                            copyParagraphInDocFooter(ph);
                        }

                    }
                    curP++;
                }
            }
        }
        // 处理完毕模板，删除文本中的模板内容
        for (int a = 0; a < templateBodySize; a++) {
            document.removeBodyElement(0);
        }
    }

    private static void copyParagraphInDocFooter(XWPFParagraph ph){
        XWPFParagraph createParagraph = document.createParagraph();
        // 设置段落样式
        createParagraph.getCTP().setPPr(ph.getCTP().getPPr());
        // 移除原始内容
        for (int pos = 0; pos < createParagraph.getRuns().size(); pos++) {
            createParagraph.removeRun(pos);
        }
        // 添加Run标签
        for (XWPFRun s : ph.getRuns()) {
            XWPFRun targetrun = createParagraph.createRun();
            copyRun(targetrun, s);
        }
    }

    private static void copyTableInDocFooter(XWPFTable templateTable){
        List<XWPFTableRow> templateTableRows = templateTable.getRows();// 获取模板表格所有行
        XWPFTable newCreateTable = document.createTable();// 创建新表格,默认一行一列
        for (int i = 0; i < templateTableRows.size(); i++) {
            XWPFTableRow newCreateRow = newCreateTable.createRow();
            copyTableRow(newCreateRow, templateTableRows.get(i));// 复制模板行文本和样式到新行
        }
        newCreateTable.removeRow(0);// 移除多出来的第一行
    }

    private static void copyTableRow(XWPFTableRow target, XWPFTableRow source) {

        int tempRowCellsize = source.getTableCells().size();// 模板行的列数
        for (int i = 0; i < tempRowCellsize - 1; i++) {
            target.addNewTableCell();// 为新添加的行添加与模板表格对应行行相同个数的单元格
        }
        // 复制样式
        target.getCtRow().setTrPr(source.getCtRow().getTrPr());
        // 复制单元格
        for (int i = 0; i < target.getTableCells().size(); i++) {
            copyTableCell(target.getCell(i), source.getCell(i));
        }
    }

    private static void copyTableCell(XWPFTableCell newTableCell, XWPFTableCell templateTableCell) {
        // 列属性
        newTableCell.getCTTc().setTcPr(templateTableCell.getCTTc().getTcPr());
        // 删除目标 targetCell 所有文本段落
        for (int pos = 0; pos < newTableCell.getParagraphs().size(); pos++) {
            newTableCell.removeParagraph(pos);
        }
        // 添加新文本段落
        for (XWPFParagraph sp : templateTableCell.getParagraphs()) {
            XWPFParagraph targetP = newTableCell.addParagraph();
            copyParagraph(targetP, sp);
        }
    }

    private static void copyParagraph(XWPFParagraph newParagraph, XWPFParagraph templateParagraph) {
        // 设置段落样式
        newParagraph.getCTP().setPPr(templateParagraph.getCTP().getPPr());
        // 添加Run标签
        for (int pos = 0; pos < newParagraph.getRuns().size(); pos++) {
            newParagraph.removeRun(pos);

        }
        for (XWPFRun s : templateParagraph.getRuns()) {
            XWPFRun targetrun = newParagraph.createRun();
            copyRun(targetrun, s);
        }

    }

    private static void copyRun(XWPFRun newRun, XWPFRun templateRun) {
        newRun.getCTR().setRPr(templateRun.getCTR().getRPr());
        // 设置文本
        newRun.setText(templateRun.text());
    }

}

