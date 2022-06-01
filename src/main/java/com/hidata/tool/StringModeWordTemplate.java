package com.hidata.tool;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import java.io.IOException;
import java.io.InputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class StringModeWordTemplate {

    private XWPFDocument document;

    public XWPFDocument getDocument() {
        return document;
    }

    public void setDocument(XWPFDocument document) {
        this.document = document;
    }

    public StringModeWordTemplate(InputStream inputStream) throws IOException {
        document = new XWPFDocument(inputStream);
    }

    public void replaceDocument(Map<String, Object> parametersMap) throws XmlException, IOException {
        List<IBodyElement> bodyElements = document.getBodyElements();  // 所有对象（段落+表格）

        int curT = 0;// 当前操作表格对象的索引
        int curP = 0;// 当前操作段落对象的索引
        for (int i = 0; i < bodyElements.size(); i++) {
            IBodyElement bodyElement = bodyElements.get(i);
            System.out.println(bodyElement.getElementType());
            if (BodyElementType.TABLE.equals(bodyElement.getElementType())) {
                // 处理表格
                List<XWPFTable> tables = bodyElement.getBody().getTables();
                // 注意这里只能这么拿table，每次处理一个，因为bodyElements.size()肯定大于表数，在下面判空即可
                XWPFTable table = tables.get(curT);
                if (table != null) {
                    String multiIndicatorCellString = "";
                    // 逐一分析每个单元格里的元素，如果包含SF_VFUN或者SF_MDFUN则要处理,获取数据之后替换原来的元素
                    List<XWPFTableRow> tableRows = table.getRows();
                    for (XWPFTableRow row : tableRows) {
                        List<XWPFTableCell> tableCells = row.getTableCells();
                        for (XWPFTableCell cell : tableCells) {
                            String cellString = cell.getText();
                            String targetCellString = "";
                            // 获取到每一行中的单元格
                            if (cellString.contains("SF_VFUN")) {
                                // 单维指标
                                targetCellString = getSFVFUNString(cell.getText());
                                cell.removeParagraph(0);
                                cell.setText(targetCellString);
                            } else if (cellString.contains("SF_MDFUN")) {
                                // 多维指标
                                // 向上传递单元格的内容
                                cell.removeParagraph(0);
                                multiIndicatorCellString = cellString;
                                break;
                            }
                        }
                        if (!"".equals(multiIndicatorCellString)){
                            // 多维指标则退出循环
                            break;
                        }
                    }
                    if (!"".equals(multiIndicatorCellString)){

                        // 使用多维指标的数据填充方式
//                        List<List<String>> multiIndicatorList = getMultiIndicatorList(multiIndicatorCellString);
                        List<List<String>> multiIndicatorList = new ArrayList<>();
//                        if (multiIndicatorList.size()==0) {
//                            curT++;
//                            continue;
//                        }
                        // 判断表头
                        int headRowNumber = 0;
                        int tableColSize = tableRows.get(tableRows.size() - 1).getTableCells().size();  //最后一行的cell数
                        for (XWPFTableRow row : tableRows) {
                            if (row.getTableCells().size() != tableColSize) {
                                headRowNumber++;
                            } else {
                                headRowNumber++;
                                break;
                            }
                        }
                        // 判断除去表头之后的第一行是否为空
                        List<XWPFTableCell> firstContentCells = tableRows.get(headRowNumber).getTableCells();
                        XWPFTableRow tempRow = tableRows.get(tableRows.size()-1);  // 最后一行作为模板行
                        if ("".equals(firstContentCells.get(0).getText())) {
                            int dataIndex = headRowNumber;  // 数据集下标
                            for (List<String> fields : multiIndicatorList) {
                                CTRow ctrow = CTRow.Factory.parse(tempRow.getCtRow().newInputStream());
                                XWPFTableRow newRow = new XWPFTableRow(ctrow, table);
                                for (int k = 0; k < fields.size(); k++) {
                                    XWPFTableCell cell = newRow.getTableCells().get(k);
                                    cell.setText(fields.get(k));  // 由于我们的需求都是空行，这个地方直接setText就行
//                                    for (XWPFParagraph p : cell.getParagraphs()) {
//                                        for (XWPFRun r : p.getRuns()) {
//                                            r.setText(fields.get(k));//要深入到原cell中的run替换内容才能保证样式一致
//                                        }
//                                    }
                                }
                                table.addRow(newRow, dataIndex + 1);
                                dataIndex++;

                            }
                            table.removeRow(headRowNumber);  // 移除模板行

                        } else {
                            // 直接填充数据
                            int rowIndex = 0;
                            for (XWPFTableRow row : table.getRows().subList(headRowNumber, table.getRows().size())) {

                                List<XWPFTableCell> tableCells = row.getTableCells();
                                int cellIndex = 0;
                                for (XWPFTableCell cell : tableCells) {

                                    if ("".equals(cell.getText())) {
                                        String cellString = "";
                                        try{
                                            cellString = multiIndicatorList.get(rowIndex).get(cellIndex);
                                        } catch (IndexOutOfBoundsException exception){
                                            cellString = "-";
                                        }
                                        cell.setText(cellString);
                                        cellIndex++;
                                    }
                                }
                                rowIndex++;
                            }
                        }
                    }
                    curT++;
                }
            }
            else if (BodyElementType.PARAGRAPH.equals(bodyElement.getElementType())) {// 处理段落
                XWPFParagraph ph = bodyElement.getBody().getParagraphArray(curP);
                if (ph != null) {
                    replaceParagraph(ph, parametersMap);
                    curP++;
                }
            }
        }
    }

    public void replaceParagraph(XWPFParagraph xWPFParagraph, Map<String, Object> parametersMap) {
        List<XWPFRun> runs = xWPFParagraph.getRuns();
        String xWPFParagraphText = xWPFParagraph.getText();
        String regEx = "\\{.+?\\}";
        Pattern pattern = Pattern.compile(regEx);
        Matcher matcher = pattern.matcher(xWPFParagraphText);//正则匹配字符串{****}

        if (matcher.find()) {
            // 查找到有标签才执行替换
            int beginRunIndex = xWPFParagraph.searchText("{", new PositionInParagraph()).getBeginRun();// 标签开始run位置
            int endRunIndex = xWPFParagraph.searchText("}", new PositionInParagraph()).getEndRun();// 结束标签
            StringBuffer key = new StringBuffer();

            if (beginRunIndex == endRunIndex) {
                // {**}在一个run标签内
                XWPFRun beginRun = runs.get(beginRunIndex);
                String beginRunText = beginRun.text();

                int beginIndex = beginRunText.indexOf("{");
                int endIndex = beginRunText.indexOf("}");
                int length = beginRunText.length();

                if (beginIndex == 0 && endIndex == length - 1) {
                    // 该run标签只有{**}
                    XWPFRun insertNewRun = xWPFParagraph.insertNewRun(beginRunIndex);
                    insertNewRun.getCTR().setRPr(beginRun.getCTR().getRPr());
                    // 设置文本
                    key.append(beginRunText.substring(1, endIndex));
                    insertNewRun.setText(getValueBykey(key.toString(),parametersMap));
                    xWPFParagraph.removeRun(beginRunIndex + 1);
                } else {
                    // 该run标签为**{**}** 或者 **{**} 或者{**}**，替换key后，还需要加上原始key前后的文本
                    XWPFRun insertNewRun = xWPFParagraph.insertNewRun(beginRunIndex);
                    insertNewRun.getCTR().setRPr(beginRun.getCTR().getRPr());
                    // 设置文本
                    key.append(beginRunText.substring(beginRunText.indexOf("{")+1, beginRunText.indexOf("}")));
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
                int beginIndex = beginRunText.indexOf("{");
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
                    String textString=beginRunText.substring(0,beginRunText.indexOf("{"))+getValueBykey(key.toString(),parametersMap);
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

    private String getSFVFUNString(String cellString){
        return "123";
    }

    private List<List<String>> getMultiIndicatorList(String cellString){
        System.out.println(cellString);
        List<List<String>> dataList = new ArrayList<>();

        ArrayList<String> list1 = new ArrayList<>();
        list1.add("1-1");
        list1.add("1-2");
        list1.add("1-3");
        list1.add("1-4");

        ArrayList<String> list2 = new ArrayList<>();
        list2.add("2-1");
        list2.add("2-2");
        list2.add("2-3");
        list2.add("2-4");

        dataList.add(list1);
        dataList.add(list2);

        return dataList;
    }

    private String getValueBykey(String key, Map<String, Object> map) {
        String returnValue="";
        if (key != null) {
            try {
                returnValue=map.get(key)!=null ? map.get(key).toString() : "";
            } catch (Exception e) {
                // TODO: handle exception
                System.out.println("key:"+key+"***"+e);
                returnValue="";
            }

        }
        return returnValue;
    }

//    public void removeDocumentByCondition(Map<String, Object> parametersMap) {
//
//        List<IBodyElement> bodyElements = document.getBodyElements();// 所有对象（段落+表格）
//        int templateBodySize = bodyElements.size();// 标记模板文件（段落+表格）总个数
//        boolean deleteFlag = false;
//        int startIndex = 0;
//
//        String regEx = "\\$(.*?)=\"(.*?)\"";
//        Pattern pattern = Pattern.compile(regEx);
//
//        HashSet<Integer> removeIndexSet = new HashSet<>();
//
//        int curP = 0;// 当前操作段落对象的索引
//        for (int a = 0; a < templateBodySize; a++) {
//
//            IBodyElement body = bodyElements.get(a);
//            if (BodyElementType.PARAGRAPH.equals(body.getElementType())) {
//                XWPFParagraph ph = body.getBody().getParagraphArray(curP);
//                if (ph != null) {
//                    String text = ph.getText();
//                    if (text.contains("IF")) {
//                        Matcher matcher = pattern.matcher(text);
//                        if (matcher.find()) {
//                            String keyString = matcher.group(1);
//                            String valueString = matcher.group(2);
//                            if (valueString.equals(parametersMap.get(keyString))) {
//                                // 符合条件，只需要删除段落本身
//                                removeIndexSet.add(a);
//                                document.removeBodyElement(document.getPosOfParagraph(ph));
//                            } else {
////                                // if不符合，要删除区间内包括本身的所有段落和表格
////                                if (text.startsWith("IF")) {
////                                    startIndex = a;
////                                } else if (text.startsWith("ENDIF")) {
////                                    for (int i = startIndex; i <= a; i++) {
////                                        removeIndexSet.add(i);
////                                    }
////                                }
//
//                            }
//                        }
//                    }
//                }
//            }
//            curP++;
//        }
//
////        // 处理完毕模板，删除文本中的模板内容
////        for (Integer integer : removeIndexSet) {
////            document.removeBodyElement(integer);
////        }
//
//    }
    private boolean isIfExists(){
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

    public void removeDocumentByCondition(Map<String, Object> parametersMap) {
        while (this.isIfExists()){
            removeDocument(parametersMap);
        }
    }

    // 循环调用直到文档中不包含含有if的段落
    public void removeDocument(Map<String, Object> parametersMap) {

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
                                        XWPFTable table = body.getBody().getTableArray(curT);
                                        if (table != null) {
                                            copyTableInDocFooter(table);
                                            curT++;
                                            a++;
                                        }
                                    } else if (BodyElementType.PARAGRAPH.equals(subElement.getElementType())){
                                        XWPFParagraph innerPh = body.getBody().getParagraphArray(curP+1);
                                        if (innerPh != null) {
                                            String innerPhText = innerPh.getText();
                                            if (innerPhText.startsWith("ENDIF("+completeString)){
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
                                if (text.startsWith("IF")) {
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
                                                if (innerPhText.startsWith("ENDIF("+completeString)){
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
                        copyParagraphInDocFooter(ph);
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

    private void copyParagraphInDocFooter(XWPFParagraph ph){
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

    private void copyTableInDocFooter(XWPFTable templateTable){
        List<XWPFTableRow> templateTableRows = templateTable.getRows();// 获取模板表格所有行
        XWPFTable newCreateTable = document.createTable();// 创建新表格,默认一行一列
        for (int i = 0; i < templateTableRows.size(); i++) {
            XWPFTableRow newCreateRow = newCreateTable.createRow();
            copyTableRow(newCreateRow, templateTableRows.get(i));// 复制模板行文本和样式到新行
        }
        newCreateTable.removeRow(0);// 移除多出来的第一行
        document.createParagraph();// 添加回车换行
    }

    private void copyTableRow(XWPFTableRow target, XWPFTableRow source) {

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

    private void copyTableCell(XWPFTableCell newTableCell, XWPFTableCell templateTableCell) {
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

    private void copyParagraph(XWPFParagraph newParagraph, XWPFParagraph templateParagraph) {
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

    private void copyRun(XWPFRun newRun, XWPFRun templateRun) {
        newRun.getCTR().setRPr(templateRun.getCTR().getRPr());
        // 设置文本
        newRun.setText(templateRun.text());
    }

}

