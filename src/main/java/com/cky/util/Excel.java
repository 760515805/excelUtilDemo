package com.cky.util;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class Excel<T> {
	/**
	 * 
	* @Title: 根据列名将集合转换成Excel
	* @Description:mapping
	* 示例:生成只有用户名，和密码两列的表
	* Map<String,String> mapping=new LinkedHashMap<String,String>();
 		mapping.put("用户名", "id");
		mapping.put("密码", "username");
	*  @param models实体类
	*  @param mapping要插入到表中的列名以及对应的实体类的属性
	*  @param headLine表的标题
	*  @param outputFile输出的文件
	*  @throws Exception    
	* @return void
	* @throws
	 */
	public void setToExcelByColumn(List<T> models,Map<String,String> mapping,String headLine,OutputStream out) throws Exception{
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
		insertColumnName(1,sheet,mapping);
		
		//*从第2行开始插入数据
		insertColumnDate(2,models,sheet,mapping);
		
		//输出表格文件
		try {
			wb.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	/**
	 * 
	* @Title: 将集合所有的数据导入表格  
	* @Description: TODO
	*  @param models
	*  @param headLine
	*  @param out    
	* @return void
	* @throws
	 */
	public void setToExcel(List<T> models,String headLine,OutputStream out){
		Map<String,String> mapping=new LinkedHashMap<String,String>();
		Iterator<T> iter=models.iterator();
		T model;
		//获取一个元素
		if(iter.hasNext()){
			model=iter.next();
			Field[] fields=model.getClass().getDeclaredFields();
			 String[]   name   =   new   String[fields.length]; 
	         Object[]   value   =   new   Object[fields.length]; 
	               
	         try{ 
               Field.setAccessible(fields,   true); 
               for   (int   i   =   0;   i   <   name.length;   i++)   { 
                      name[i]   =   fields[i].getName(); 
                      System.out.println(name[i]   +   "-> "); 
                      value[i]   =   fields[i].get(model); 
                      System.out.println(value[i]); 
                      mapping.put(isNull(name[i]).toString(), isNull(value[i]).toString());
               } 
               setToExcelByColumn(models, mapping, headLine, out);
	         } 
	         catch(Exception   e){ 
	                  e.printStackTrace(); 
	         } 
	         
		}
		
	}
	
	
	/**
	 * 此方法作用是创建表头的列名
	 * @param mapping	要创建的表的列名与实体类的属性名的映射集合
	 * @param row		指定行创建列名
	 * @return
	 */
	private  void insertColumnName(int rowNum,HSSFSheet sheet,Map<String,String> mapping){
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
	
	private void insertColumnDate(int rowNum,List<T> models,HSSFSheet sheet,Map<String,String> mapping) throws Exception{
		for (T model : models) {
			//创建新的一行
			HSSFRow rowTemp =sheet.createRow(rowNum++);
			System.out.println("创建了第"+rowNum+"行");
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
				System.out.print("创建了第"+columnNum+"个格子");
				HSSFCell cellTemp=rowTemp.createCell(columnNum++);
				//将此属性值放入格子
				System.out.println(",放入了"+isNull(obj).toString());
				cellTemp.setCellValue(isNull(obj).toString());
			}
			System.out.println("");
		}
	}
	
	
	
	//将列名对应的属性名换成get方法名
	private  void attrToSetMethod(Entry entry){
		//获取entry的值
		
		String attrName=(String) isNull(entry.getValue());
		//连接字符串
		String getAttrMethodName="get"+captureName(attrName);
		//将方法名放入entry的值中
		entry.setValue(getAttrMethodName);
	}
	
	//首字母大写
	private  String captureName(String name) {
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
	private Object isNull(Object object){
		if(object!=null){
			return object;
		}else{
			return "";
		}
	}
	
}
