package com.hidata.tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Test {

	public static void main(String[] args) throws IOException {

		Map<String, Object> wordDataMap = new HashMap<String, Object>();// 存储报表全部数据
		Map<String, Object> parametersMap = new HashMap<String, Object>();// 存储报表中不循环的数据



		List<Map<String, Object>> table1 = new ArrayList<Map<String, Object>>();
		Map<String, Object> map1=new HashMap<>();
		map1.put("name", "张三");
		map1.put("age", "23");
		map1.put("email", "12121@qq.com");

		Map<String, Object> map2=new HashMap<>();
		map2.put("name", "李四");
		map2.put("age", "45");
		map2.put("email", "45445@qq.com");

		Map<String, Object> map3=new HashMap<>();
		map3.put("name", "Tom");
		map3.put("age", "34");
		map3.put("email", "6767@qq.com");

		table1.add(map1);
		table1.add(map2);
		table1.add(map3);




		List<Map<String, Object>> table2 = new ArrayList<Map<String, Object>>();
		Map<String, Object> map4=new HashMap<>();
		map4.put("name", "tom");
		map4.put("number", "sd1234");
		map4.put("address", "上海");

		Map<String, Object> map5=new HashMap<>();
		map5.put("name", "seven");
		map5.put("number", "sd15678");
		map5.put("address", "北京");

		Map<String, Object> map6=new HashMap<>();
		map6.put("name", "lisa");
		map6.put("number", "sd9078");
		map6.put("address", "广州");

		table2.add(map4);
		table2.add(map5);
		table2.add(map6);

		List<Map<String, Object>> table3 = new ArrayList<Map<String, Object>>();
		Map<String, Object> map7=new HashMap<>();
		map7.put("asset_type", "现金及银行存款");
		map7.put("a_money", "123");
		map7.put("a_bili", "10%");
		Map<String, Object> map8=new HashMap<>();
		map8.put("asset_type", "现金及银行存款");
		map8.put("a_money", "123");
		map8.put("a_bili", "10%");
		table3.add(map7);
		table3.add(map8);



		parametersMap.put("userName", "JUVENILESS");
		parametersMap.put("time", "2022-03-24");
		parametersMap.put("sum", "3");
		parametersMap.put("year", "2022");
		parametersMap.put("PROD_CODE", "123");


		wordDataMap.put("table1", table1);
		wordDataMap.put("table2", table2);
		wordDataMap.put("table3", table3);
		wordDataMap.put("parametersMap", parametersMap);

//		File file = new File("E:\\workspace\\project\\playground\\wordTemplate\\doc\\模板.docx");//改成你本地文件所在目录
		File file = new File("E:\\workspace\\project\\playground\\wordTemplate\\doc\\demoFile.docx");//改成你本地文件所在目录


		// 读取word模板
		FileInputStream fileInputStream = new FileInputStream(file);
		WordTemplate template = new WordTemplate(fileInputStream);

		// 替换数据
		template.replaceDocument(wordDataMap);


		//生成文件
//		File outputFile=new File("E:\\workspace\\project\\playground\\wordTemplate\\doc\\输出.docx");//改成你本地文件所在目录
		File outputFile=new File("E:\\workspace\\project\\playground\\wordTemplate\\doc\\demoFileOut.docx");//改成你本地文件所在目录
		FileOutputStream fos  = new FileOutputStream(outputFile);
		template.getDocument().write(fos);

	}

}
