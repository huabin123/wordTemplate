package com.hidata.tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * 测试不用{}占位的例子，demoFile.docx
 */
public class TestSimple {

	public static void main(String[] args) throws IOException {

		Map<String, Object> wordDataMap = genWordData();  // 生成填充word需要的数据

		File file = new File("E:\\workspace\\project\\playground\\wordTemplate\\doc\\demoFile.docx");//改成你本地文件所在目录

		// 读取word模板
		FileInputStream fileInputStream = new FileInputStream(file);
		WordTemplateSimple template = new WordTemplateSimple(fileInputStream);

		// 替换数据
		template.replaceDocument(wordDataMap);

		//生成文件
		File outputFile=new File("E:\\workspace\\project\\playground\\wordTemplate\\doc\\demoFileOut.docx");//改成你本地文件所在目录
		FileOutputStream fos  = new FileOutputStream(outputFile);
		template.getDocument().write(fos);
	}

	private static Map<String, Object> genWordData(){
		Map<String, Object> wordDataMap = new HashMap<String, Object>();// 存储报表全部数据
		Map<String, Object> parametersMap = new HashMap<String, Object>();// 存储报表中不循环的数据

		List<String> singleTable1 = new ArrayList<>();
		singleTable1.add("1");
		singleTable1.add("2");
		singleTable1.add("3");

		parametersMap.put("PROD_CODE", "123");

		wordDataMap.put("singleTable1", singleTable1);
		wordDataMap.put("parametersMap", parametersMap);

		return wordDataMap;
	}
}
