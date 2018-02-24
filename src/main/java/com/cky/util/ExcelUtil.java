package com.cky.util;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ExcelUtil {
	private static final Logger logger=LoggerFactory.getLogger(ExcelUtil.class);
	/**
	 * 将对象数组转换成excel
	 * @param pojoList 	对象数组
	 * @param out		输出流
	 * @param alias		指定对象属性别名，生成列名和列顺序
	 * @param headLine	表标题
	 * @throws Exception 
	 */
	public static void pojo2Excel(List pojoList,OutputStream out,LinkedHashMap<String,String> alias,String headLine) throws Exception {
		//创建一个工作本
		HSSFWorkbook wb=new HSSFWorkbook();
		//创建一个表
		HSSFSheet sheet=wb.createSheet();
		//创建第一行
		HSSFRow	  row=sheet.createRow(0);
		HSSFCell cell=row.createCell(0);
		cell.setCellValue(headLine);
		
		sheet.addMergedRegion(new CellRangeAddress(0,0,0,3));
		
		//在第一行插入列名
		insertColumnName(1,sheet,alias);
		
		//*从第2行开始插入数据
		insertColumnDate(2,pojoList,sheet,alias);
		
		//输出表格文件
		try {
			wb.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	/**
	 * 将对象数组转换成excel
	 * @param pojoList 	对象数组
	 * @param out		输出流
	 * @param alias		指定对象属性别名，生成列名和列顺序
	 * @throws Exception 
	 */
	public static void pojo2Excel(List pojoList,OutputStream out,LinkedHashMap<String,String> alias) throws Exception {
		//获取类名作为标题
		String headLine="";
		if(pojoList.size()>0) {
			Object pojo=pojoList.get(0);
			Class<? extends Object> claz = pojo.getClass();
			headLine = claz.getName();
			pojo2Excel(pojoList, out, alias, headLine);
		}
	}
	/**
	 * 将对象数组转换成excel,列名为对象属性名
	 * @param pojoList 	对象数组
	 * @param out		输出流
	 * @param headLine	表标题
	 * @throws Exception 
	 */
	public static void pojo2Excel(List pojoList,OutputStream out,String headLine) throws Exception{
		//获取类的属性作为列名
		LinkedHashMap<String,String> alias=new LinkedHashMap<String,String>();
		if(pojoList.size()>0) {
			Object pojo = pojoList.get(0);
			Field[] fields=pojo.getClass().getDeclaredFields();
			String[]   name   =   new   String[fields.length]; 
			Field.setAccessible(fields,   true); 
            for   (int   i   =   0;   i   <   name.length;   i++)   { 
                   name[i]   =   fields[i].getName(); 
                   alias.put(isNull(name[i]).toString(), isNull(name[i]).toString());
            } 
            pojo2Excel(pojoList,out,alias,headLine);
		}
	}
	/**
	 * 将对象数组转换成excel，列名默认为对象属性名，标题为类名
	 * @param pojoList	对象数组
	 * @param out		输出流
	 * @throws Exception 
	 */
	public static void pojo2Excel(List pojoList,OutputStream out) throws Exception{
		//获取类的属性作为列名
		LinkedHashMap<String,String> alias=new LinkedHashMap<String,String>();
		//获取类名作为标题
		String headLine="";
		if(pojoList.size()>0) {
			Object pojo = pojoList.get(0);
			Class<? extends Object> claz = pojo.getClass();
			headLine=claz.getName();
			Field[] fields=claz.getDeclaredFields();
			String[]   name   =   new   String[fields.length]; 
			Field.setAccessible(fields,   true); 
            for   (int   i   =   0;   i   <   name.length;   i++)   { 
                   name[i]   =   fields[i].getName(); 
                   alias.put(isNull(name[i]).toString(), isNull(name[i]).toString());
            } 
            pojo2Excel(pojoList,out,alias,headLine);
		}
	}
	/**
	 * 将excel表转换成指定类型的对象数组
	 * @param claz 	类型
	 * @param alias	列别名
	 * @return
	 * @throws IOException 
	 * @throws IllegalArgumentException 
	 * @throws IllegalAccessException 
	 * @throws SecurityException 
	 * @throws NoSuchFieldException 
	 * @throws InstantiationException 
	 * @throws InvocationTargetException 
	 */
	public static<T>List<T> excel2Pojo(InputStream inputStream,Class<T> claz,LinkedHashMap<String,String> alias) throws IOException, IllegalArgumentException, IllegalAccessException, NoSuchFieldException, SecurityException, InstantiationException, InvocationTargetException{
		HSSFWorkbook wb = new HSSFWorkbook(inputStream);
		HSSFSheet sheet = wb.getSheetAt(0);
		
		//获取列信息，Map<类属性名，对应一行的第几列>
		Map<String,Integer> propertyMap=new HashMap<>();
		
		HSSFRow propertyRow = sheet.getRow(1);
		short firstCellNum = propertyRow.getFirstCellNum();
		short lastCellNum = propertyRow.getLastCellNum();
		
		for(int i=firstCellNum;i<lastCellNum;i++) {
			Cell cell = propertyRow.getCell(i);
			if(cell==null) {
				continue;
			}
			//列名
			String cellValue = cell.getStringCellValue();
			//对应属性名
			String propertyName = alias.get(cellValue);
			propertyMap.put(propertyName, i);
		}
		//对象数组
		List<T> pojoList=new ArrayList<>();
		for (Row row : sheet) {
			//跳过第一行标题
			if(row.getRowNum()<3) {
				continue;
			}
			T instance = claz.newInstance();
			Set<Entry<String, Integer>> entrySet = propertyMap.entrySet();
			for (Entry<String, Integer> entry : entrySet) {
				BeanUtils.setProperty(instance, entry.getKey(), row.getCell(entry.getValue()).getStringCellValue().toString());
			}
			pojoList.add(instance);
		}
		return pojoList;
	}
	/**
	 * 将excel表转换成指定类型的对象数组，列名即作为对象属性
	 * @param claz 	类型
	 * @return
	 * @throws IOException 
	 * @throws InstantiationException 
	 * @throws SecurityException 
	 * @throws NoSuchFieldException 
	 * @throws IllegalAccessException 
	 * @throws IllegalArgumentException 
	 * @throws InvocationTargetException 
	 */
	public static<T>List<T> excel2Pojo(InputStream inputStream,Class<T> claz) throws IllegalArgumentException, IllegalAccessException, NoSuchFieldException, SecurityException, InstantiationException, IOException, InvocationTargetException{
		LinkedHashMap<String,String> alias = new LinkedHashMap<String, String>();
		Field[] fields = claz.getDeclaredFields();
		for (Field field : fields) {
			alias.put(field.getName(), field.getName());
		}
		List<T> pojoList = excel2Pojo(inputStream, claz, alias);
		return pojoList;
	}
	
	/**
	 * 此方法作用是创建表头的列名
	 * @param mapping	要创建的表的列名与实体类的属性名的映射集合
	 * @param row		指定行创建列名
	 * @return
	 */
	private static void insertColumnName(int rowNum,HSSFSheet sheet,Map<String,String> mapping){
		HSSFRow row =sheet.createRow(rowNum);
		//列的数量
		int columnCount=0;
		Iterator columnIter=mapping.entrySet().iterator();
		//在第一行创建列名
		while(columnIter.hasNext()){
			//获取一个mapping中第一个键值对entry
			Map.Entry entry=(Entry) columnIter.next();
			
			//将entry中的值由属性名换成“get+属性”的方法名，供之后获取method
			attrToSetMethod(entry);
			
			//创建第一行的第columnCount个格子
			HSSFCell cell1_0=row.createCell(columnCount++);
			
			//将此格子的值设置为mapping中的键名
			cell1_0.setCellValue(isNull(entry.getKey()).toString());
		}
	}
	
	private static <T>void insertColumnDate(int rowNum,List<T> models,HSSFSheet sheet,Map<String,String> mapping) throws Exception{
		for (T model : models) {
			//创建新的一行
			HSSFRow rowTemp =sheet.createRow(rowNum++);
			logger.info("创建了第:{}行", rowNum);
			//获取列的迭代
			Iterator methodNameIter=mapping.entrySet().iterator();
			//从第0个格子开始创建
			int columnNum=0;
			
			while(methodNameIter.hasNext()){
				Map.Entry<String, String> methodNameEntry=(Entry<String, String>) methodNameIter.next();
				String methodName=methodNameEntry.getValue();
				//获取此列对应实体类的get方法
				Method method=model.getClass().getMethod(methodName);
				//调用此方法获取实体类的属性值
				Object obj=method.invoke(model);
		
				//创建一个格子
				HSSFCell cellTemp=rowTemp.createCell(columnNum++);
				logger.info("创建了第：{}个格子，存入：{}",columnNum,isNull(obj).toString());
				//将此属性值放入格子
				cellTemp.setCellValue(isNull(obj).toString());
			}
		}
	}
	
	
	
	//将列名对应的属性名换成get方法名
	private static void attrToSetMethod(Entry entry){
		//获取entry的值
		
		String attrName=(String) isNull(entry.getValue());
		//连接字符串
		String getAttrMethodName="get"+captureName(attrName);
		//将方法名放入entry的值中
		entry.setValue(getAttrMethodName);
	}
	
	//首字母大写
	private static String captureName(String name) {
			if(name==""||name==null){
				return "";
			}
		        char[] cs=name.toCharArray();
		        if(cs[0]<97||cs[0]>122){
		        	//首字母不在小写范围,直接返回原值
		        	return String.valueOf(cs);
		        }
		        //首字母-32,变成大写
		        cs[0]-=32;
		        return String.valueOf(cs);
	 }
	//判断是否为空，若为空设为""
	private static Object isNull(Object object){
		if(object!=null){
			return object;
		}else{
			return "";
		}
	}
}
