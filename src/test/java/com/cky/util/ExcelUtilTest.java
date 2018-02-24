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

import com.cky.util.ExcelUtil;
import com.cky.util.User;

public class ExcelUtilTest{
	private static final Logger logger=LoggerFactory.getLogger(ExcelUtilTest.class);
	@Test
	public void pojo2Excel1() throws Exception {
		File file = new File("C:\\Users\\chenkeyu\\Work\\1.xls");
		FileOutputStream fos = new FileOutputStream(file);
		
		//对象集合
		List<User> pojoList=new ArrayList<>();
		for(int i=0;i<5;i++) {
			User user = new User();
			user.setName("老王");
			user.setAge(11);
			pojoList.add(user);
		}
		//别名
		LinkedHashMap<String, String> alias = new LinkedHashMap<>();
		alias.put("姓名", "name");
		alias.put("年龄","age");
		//标题
		String headLine="用户";
		
		ExcelUtil.pojo2Excel(pojoList, fos, alias, headLine);
	}
	@Test 
	public void excel2Pojo() throws Exception {
		FileInputStream fis = new FileInputStream("C:\\Users\\chenkeyu\\Work\\1.xls");
		LinkedHashMap<String, String> alias = new LinkedHashMap<>();
		alias.put("姓名","name");
		alias.put("年龄","age");
		List<User> pojoList = ExcelUtil.excel2Pojo(fis, User.class, alias);
		logger.info(pojoList.toString());
	}
}
