package com.cky.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.cky.bean.User;
import com.cky.util.ExcelUtil;

public class ExcelUtilTest{
	private static final Logger logger=LoggerFactory.getLogger(ExcelUtilTest.class);
	@Test
	/**
	 * 对象转换成excel文件测试
	 * @throws Exception
	 */
	public void pojo2Excel1() throws Exception {
		//将生成的excel转换成文件，还可以用作文件下载
		File file = new File("C:\\Users\\chenkeyu\\Work\\1.xls");
		FileOutputStream fos = new FileOutputStream(file);
		
		//对象集合
		List<User> pojoList=new ArrayList<>();
		for(int i=0;i<5;i++) {
			User user = new User();
			user.setName("老李");
			user.setAge(50);
			pojoList.add(user);
		}
		//设置属性别名（列名）
		LinkedHashMap<String, String> alias = new LinkedHashMap<>();
		alias.put("name", "姓名");
		alias.put("age","年龄");
		//标题
		String headLine="用户表";
		
		ExcelUtil.pojo2Excel(pojoList, fos, alias, headLine);
	}
	@Test 
	/**
	 * excel文件转换成对象测试
	 * @throws Exception
	 */
	public void excel2Pojo() throws Exception {
		//指定输入文件
		FileInputStream fis = new FileInputStream("C:\\Users\\chenkeyu\\Work\\1.xls");
		//指定每列对应的类属性
		LinkedHashMap<String, String> alias = new LinkedHashMap<>();
		alias.put("姓名","name");
		alias.put("年龄","age");
		//转换成指定类型的对象数组
		List<User> pojoList = ExcelUtil.excel2Pojo(fis, User.class, alias);
		logger.info(pojoList.toString());
	}
}
