package com.hidata;

import com.hidata.tool.StringModeWordTemplate;
import com.hidata.tool.WordTemplate;
import org.apache.xmlbeans.XmlException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class StringModeTest {

	public static void main(String[] args) throws IOException, XmlException {

		Map<String, Object> parametersMap = new HashMap<String, Object>();// 存储报表中不循环的数据

		parametersMap.put("PROD_CODE", "001");
		parametersMap.put("REPORT_TYPE", "1");
		parametersMap.put("END_DATE", "31");

		File file = new File("E:\\workspace\\project\\playground\\wordTemplate\\doc\\StringModeSourceFile.docx");//改成你本地文件所在目录

		// 读取word模板
		FileInputStream fileInputStream = new FileInputStream(file);
		StringModeWordTemplate template = new StringModeWordTemplate(fileInputStream);

		// 根据条件移除不需要的表格段落
		template.removeDocumentByCondition(parametersMap);

		// 替换数据
		template.replaceDocument(parametersMap);

		//生成文件
		File outputFile=new File("E:\\workspace\\project\\playground\\wordTemplate\\doc\\StringModeOutFile.docx");//改成你本地文件所在目录
		FileOutputStream fos  = new FileOutputStream(outputFile);
		template.getDocument().write(fos);
	}

}
