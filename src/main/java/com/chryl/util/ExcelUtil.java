package com.chryl.util;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

public class ExcelUtil {

    /***
     * Excel数据转类对象
     * @param clazz 类
     * @param is Excel文件流
     * @param fileType 文件类型
     * @param columnIndexMap 每个属性对应的列
     * @return {"businessName":1},{"address":2},{"belongOrg":3};
     * @throws IOException
     */
    public static <T> ArrayList<T> excle2Object(Class<T> clazz, InputStream is, String fileType, Map<String, Integer> columnIndexMap) throws Exception {
        ArrayList<T> arrayList = new ArrayList<T>();//创建空的集合用于存储数据
        Workbook wb = null;//声明一个workbook对象
        try {
            if (fileType.equals("xlsx")) {
                wb = new XSSFWorkbook(is);
            } else {
                wb = new HSSFWorkbook(is);
            }
            Sheet sheet = wb.getSheetAt(0);//获取sheet对象
            Field[] fi = clazz.getDeclaredFields();//获取类属性信息
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                T obj = clazz.newInstance();
                System.out.println("第" + i + " 行数据");

                for (int j = 0; j < fi.length; j++) {

                    Field field = fi[j];
                    String fieldName = field.getName();

                    if (columnIndexMap.containsKey(fieldName)) {
                        int valueIndex = columnIndexMap.get(fieldName);
                        Cell cell = row.getCell(valueIndex);
                        if (cell == null) {
                            throw new RuntimeException("字段: " + fieldName + " 序号:" + valueIndex + ":与模板不一致!");
                        }
                        field.setAccessible(true);
                        String methodName = "set" + field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1);
                        Method method = obj.getClass().getMethod(methodName, field.getType());
                        //在excel中获取的值都是String,所以要根据属性的类型进行值的转换
                        System.out.println("methodName: " + methodName + " fieldType: " + field.getType().toString());
                        DecimalFormat df = new DecimalFormat("###.####");
                        System.out.println(field.getType().toString());
                        switch (field.getType().toString()) {
                            case "class java.lang.String":
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_NUMERIC: // 数字
                                        //日期字符串转换
                                        short dataFormat = cell.getCellStyle().getDataFormat();
                                        if (14 == dataFormat || 31 == dataFormat) {// yyyy-MM-dd 14 或 yyyy-MM-dd HH:mm:ss 31
                                            double dateVal = cell.getNumericCellValue();//dateDouble
                                            Date javaDate = DateUtil.getJavaDate(dateVal);//javaDate
                                            String yyyyMMdd = sdf.format(javaDate);
                                            method.invoke(obj, yyyyMMdd);
                                            break;
                                        }
                                        method.invoke(obj, String.valueOf(df.format(cell.getNumericCellValue())));
                                        break;
                                    case Cell.CELL_TYPE_STRING: // 字符串
                                        method.invoke(obj, cell.getStringCellValue());
                                        break;
                                }
                                break;
                            case "class java.math.BigDecimal":
                                System.out.println(cell.getCellType());
                                System.out.println(cell.getCellTypeEnum());
                                switch (cell.getCellType()) {
                                    case 0:
                                        method.invoke(obj, new BigDecimal(CommonUtil.getTwoDecimal((cell.getNumericCellValue()))));
                                        break;
                                    case 1:
                                        System.out.println(cell.getStringCellValue());
                                        method.invoke(obj, new BigDecimal(cell.getStringCellValue()));
                                        break;
                                    case 3:
                                        method.invoke(obj, new BigDecimal("0.00"));
                                        break;
                                }
                                break;
                            case "int":
                            case "Integer":
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_NUMERIC: // 数字
                                        method.invoke(obj, String.valueOf(df.format(cell.getNumericCellValue())));
                                        break;
                                    case Cell.CELL_TYPE_STRING: // 字符串
                                        method.invoke(obj, cell.getStringCellValue());
                                        break;
                                }
                                break;
                            case "class java.util.Date":
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_STRING: // 字符串
                                        method.invoke(obj, sdf.parse(cell.getStringCellValue()));
                                        break;
                                }
                                break;
                            default:

                        }
                    }
                }
                arrayList.add(obj);
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            wb.close();
        }
        return arrayList;

    }

    /**
     * 将excel中的数据按行转换为object对象
     *
     * @param fileType
     * @param is
     * @param keys
     * @return
     * @throws Exception
     */
    public static ArrayList<Map<String, Object>> excelToObject(String fileType, InputStream is, String[] keys) throws Exception {
        ArrayList<Map<String, Object>> arrayList = new ArrayList<Map<String, Object>>();//创建空的集合用于存储数据
        Workbook wb = null;//声明一个workbook对象
        try {
            if (fileType.equals("xlsx")) {
                wb = new XSSFWorkbook(is);
            } else {
                wb = new HSSFWorkbook(is);
            }
            Sheet sheet = wb.getSheetAt(0); //获取sheet对象
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (null == row) continue;
                Map<String, Object> map = new HashMap<String, Object>();
                for (int j = 0; j < keys.length; j++) {
                    switch (row.getCell(j).getCellTypeEnum()) {
                        case NUMERIC: // 数字   :
                            map.put(keys[j], row.getCell(j).getNumericCellValue());
                            break;
                        case STRING: // 字符串
                            map.put(keys[j], row.getCell(j).getStringCellValue());
                            break;
                        case BLANK://空
                            map.put(keys[j], "");
                            break;
                    }

                }
                arrayList.add(map);
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            wb.close();
        }
        return arrayList;
    }


    /**
     * 将excel文件中的内容转化为类对象
     *
     * @param clazz    类
     * @param is       Excel文件流
     * @param fileType 文件类型
     *                 //@param columnIndexMap   模板中字段名对应索引的Map
     * @return list--list中存放excel内容中的类对象
     * @throws Exception
     */
    public static <T> ArrayList<T> excle2Object1(Class<T> clazz, InputStream is, String fileType) throws Exception {
        ArrayList<T> arrayList = new ArrayList<T>();//创建空的集合用于存储数据
        Workbook wb = null;//声明一个workbook对象
        try {
            if (fileType.equals("xlsx")) {
                wb = new XSSFWorkbook(is);
            } else {
                wb = new HSSFWorkbook(is);
            }
            Sheet sheet = wb.getSheetAt(0);//获取sheet对象
            Field[] fi = clazz.getDeclaredFields();//获取类属性信息
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

            T obj = clazz.newInstance();//创建类的实例

            for (int i = 1; i <= sheet.getLastRowNum(); i++) {//从1开始，0为标题行
                Row row = sheet.getRow(i);//获取一行数据

                for (int j = 1; j < fi.length; j++) {//遍历类属性信息，但是 获取索引值从1开始，0为主键，不用手动录入，插入时自动生成
                    /*
                     * 	类：clazz------》fi  类属性信息
            		 * 	excel中数据：sheet对象----》row为每行的实体对象
            		 */
                    Field field = fi[j];
                    String fieldName = field.getName();
                    Cell cell = row.getCell(j - 1);
                    if ("orgcode".equals(fieldName) && (cell == null)) {
                        throw new Exception("机构号不能为空！");
                    }
                    if ("empcode".equals(fieldName) && (cell == null)) {
                        throw new Exception("揽存号不能为空！");
                    }
                    String methodName = "set" + field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1);
                    Method method = obj.getClass().getMethod(methodName, field.getType());
                    //在excel中获取的值都是String,所以要根据属性的类型进行值的转换
                    System.out.println("methodName: " + methodName + " fieldType: " + field.getType().toString());
                    if (cell == null) {
                        method.invoke(obj, cell);
                    } else {
                        field.setAccessible(true);
                        DecimalFormat df = new DecimalFormat("0");
                        switch (field.getType().toString()) {
                            case "class java.lang.String":
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_NUMERIC: // 数字
                                        //	switch(HSSFDateUtil.isCellDateFormatted(cell)){}
                                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                            Date d = cell.getDateCellValue();
                                            method.invoke(obj, sdf.format(d));
                                            break;
                                        }
                                        method.invoke(obj, String.valueOf(df.format(cell.getNumericCellValue())));
                                        break;
                                    case Cell.CELL_TYPE_STRING: // 字符串
                                        method.invoke(obj, cell.getStringCellValue());
                                        break;
                                }
                                break;
                            case "int":
                            case "Integer":
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_NUMERIC: // 数字
                                        method.invoke(obj, String.valueOf(df.format(cell.getNumericCellValue())));
                                        break;
                                    case Cell.CELL_TYPE_STRING: // 字符串
                                        method.invoke(obj, cell.getStringCellValue());
                                        break;
                                }
                                break;
                            case "class java.util.Date":
                                switch (cell.getCellType()) {
                                    case Cell.CELL_TYPE_STRING: // 字符串
                                        method.invoke(obj, sdf.parse(cell.getStringCellValue()));
                                        break;
                                }
                                break;
                            default:
                        }
                    }

                }
                arrayList.add(obj);
            }
        } catch (Exception e) {
            e.printStackTrace();
            throw e;
        } finally {
            wb.close();
        }
        return arrayList;

    }

//    public static void main(String[] args) throws Exception {
////        ArrayList<Book> arrayList = new ArrayList<Book>();
////        Book book = new Book();
////        book.setId(1);
////        book.setName("Java语言");
////        book.setType("面向对象");
////        Book book1 = new Book();
////        book1.setId(2);
////        book1.setName("西游记");
////        book1.setType("故事");
////        Book book2 = new Book();
////        book2.setId(3);
////        book2.setName("高数");
////        book2.setType("难");
////        arrayList.add(book);
////        arrayList.add(book1);
////        arrayList.add(book2);
////        ExcelUtil.excleOut(arrayList,"D:/Sourcecode/Java/fanshe/book1.xls");
//
//    	String columnNames[]={"商户名称","地址","客户所属机构","客户创建日期","客户类型","性别","证件号码","电话","个人资产","RANK"};//列名
//        String keys[]    =  {"businessName","address","belongOrg","createDate","customerType","gender","idCard","phone","privateAsset","rank"};
//
//        Map<String,Integer> columnIndexMap = new HashMap<String,Integer>();
//        for (int i = 0; i < keys.length; i++) {
//        	columnIndexMap.put(keys[i], i);
//		}
//
//
//        //导入的数据位Book类型
//    	File file = new File("D:/客户信息管理.xlsx");
//    	FileInputStream is = new FileInputStream(file);
//        ArrayList<CrmCustomer> list = ExcelUtil.excle2Object(CrmCustomer.class,is,"xlsx",columnIndexMap);
//
//       for(CrmCustomer bo:list){
//           System.out.println(bo.getIdCard()+" "+bo.getAddress()+" "+bo.getCreateDate());
//       }
//    }

}