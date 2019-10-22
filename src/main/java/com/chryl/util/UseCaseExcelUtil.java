package com.chryl.util;

import java.io.InputStream;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import com.chryl.po.ImportModel;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



@SuppressWarnings({ "rawtypes", "deprecation" })
public class UseCaseExcelUtil
{
	final static String notnullerror = "请填入第{0}行的{1},{2}不能为空";
	final static String errormsg = "第{0}行的{1}数据导入错误";
	
	public static List<Map<String, Object>> createExcelRecord(List list) {
        List<Map<String, Object>> listmap = new ArrayList<Map<String, Object>>();
        Object t = null;
        for (int j = 0; j < list.size(); j++) {
        	t = list.get(j);
        	if(t instanceof Map) {
        		listmap.add((Map<String, Object>)t);
        	}else {
        		Map<String, Object> mapValue = BeanUtil.transBean2Map(t);
                listmap.add(mapValue);
        	}
           
        }
        return listmap;
    }
	 public static Workbook createWorkBook2007(List<Map<String, Object>> list,String[] keys,String[] columnNames) {
	    	XSSFWorkbook workbook = new XSSFWorkbook();//创建一个Excel文件
	    	XSSFSheet sheet = workbook.createSheet();// 生成一个表格
           
	    	// 设置表格默认列宽度为15个字节
			//sheet.setDefaultColumnWidth((short) 15);
	    	
			// 生成一个样式
	    	XSSFCellStyle style = workbook.createCellStyle();
			// 设置这些样式
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			style.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.THIN);
			style.setAlignment(HorizontalAlignment.CENTER);
			// 生成一个字体
			XSSFFont font = workbook.createFont();
			font.setBold(true);
			// 把字体应用到当前的样式
			style.setFont(font);
			
			XSSFCellStyle contextStyle = workbook.createCellStyle();
			contextStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			contextStyle.setFillForegroundColor(HSSFColor.WHITE.index);
			contextStyle.setBorderBottom(BorderStyle.THIN);
			contextStyle.setBorderLeft(BorderStyle.THIN);
			contextStyle.setBorderRight(BorderStyle.THIN);
			contextStyle.setBorderTop(BorderStyle.THIN);
			contextStyle.setAlignment(HorizontalAlignment.CENTER);
			contextStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			int pixel = 536;
			// 产生表格标题行
			int rowIndex = 0;
			XSSFRow row = sheet.createRow(rowIndex);
			for (short columIndex = 0; columIndex < columnNames.length; columIndex++)
			{
				
				XSSFCell cell = row.createCell(columIndex);
				cell.setCellStyle(style);
				String title = columnNames[columIndex];
				cell.setCellValue(columnNames[columIndex]);
				
				//根据标题设置列宽
				int width = title.length() * pixel;
				sheet.setColumnWidth(columIndex, width < 5000 ? 5000 : width );
			}
	    	// 遍历集合数据，产生数据行 第一行存放
			if(list != null) {
				for(int i = 0 ; i < list.size(); i++ ) {
		    		rowIndex++;
		    	    XSSFRow contextRow = sheet.createRow(rowIndex);
		    		
		    		//遍历列 
		    		if(keys != null) {
		    			for(int j=0;j<keys.length;j++){
			    			
			    			Object val = list.get(i).get(keys[j]);
			    			XSSFCell contentCell = contextRow.createCell(j);
			    			
			    			Boolean isNum = false;//val 是否为数值型
			    			Boolean isInteger = false; //val 是否为整数
			    			Boolean isPercent = false; //val 是否为百分数
			    			
			    			if(val != null) {
			    				//判断data是否为数值型
			                    isNum = val.toString().matches("^(-?\\d+)(\\.\\d+)?$");
			                    //判断data是否为整数（小数部分是否为0）
			                    isInteger=val.toString().matches("^[-\\+]?[\\d]*$");
			                    //判断data是否为百分数（是否包含“%”）
			                    isPercent=val.toString().contains("%");
			    			} else {
			    				val = " ";
			    			}
			    			 XSSFDataFormat df = workbook.createDataFormat();
			                // 设置单元格内容为字符型
			                contentCell.setCellValue(val.toString());
			                contentCell.setCellStyle(contextStyle);
			                
			                // 根据内容设置列宽度
			                int columWidth = val.toString().length() * pixel;
			                int w = sheet.getColumnWidth(j);
			                if(columWidth > w && columWidth < (20 * pixel)) {
			                	sheet.setColumnWidth(j , columWidth);
			                }
			    		}
		    		}
		    	}
			}
	    	return workbook;
	    }
	
	
    public static Workbook createWorkBook(List<Map<String, Object>> list,String[] keys,String[] columnNames) {
    	HSSFWorkbook workbook = new HSSFWorkbook();//创建一个Excel文件
    	HSSFSheet sheet = workbook.createSheet();// 生成一个表格

    	// 设置表格默认列宽度为15个字节
		//sheet.setDefaultColumnWidth((short) 15);
    	
		// 生成一个样式
		HSSFCellStyle style = workbook.createCellStyle();
		// 设置这些样式
		//style.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
		style.setBorderBottom(BorderStyle.THIN);
		style.setBorderLeft(BorderStyle.THIN);
		style.setBorderRight(BorderStyle.THIN);
		style.setBorderTop(BorderStyle.THIN);
		style.setAlignment(HorizontalAlignment.CENTER);
		// 生成一个字体
		HSSFFont font = workbook.createFont();
		//font.setColor(HSSFColor.VIOLET.index);
		//font.setFontHeightInPoints((short) 12);
		font.setBold(true);
		// 把字体应用到当前的样式
		style.setFont(font);
		
		HSSFCellStyle contextStyle = workbook.createCellStyle();
//		contextStyle.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
		
		contextStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		contextStyle.setFillForegroundColor(HSSFColor.WHITE.index);
		contextStyle.setBorderBottom(BorderStyle.THIN);
		contextStyle.setBorderLeft(BorderStyle.THIN);
		contextStyle.setBorderRight(BorderStyle.THIN);
		contextStyle.setBorderTop(BorderStyle.THIN);
		contextStyle.setAlignment(HorizontalAlignment.CENTER);
		contextStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		
		int pixel = 536;
		// 产生表格标题行
		int rowIndex = 0;
		HSSFRow row = sheet.createRow(rowIndex);
		for (short columIndex = 0; columIndex < columnNames.length; columIndex++)
		{
			
			HSSFCell cell = row.createCell(columIndex);
			cell.setCellStyle(style);
			String title = columnNames[columIndex];
			cell.setCellValue(columnNames[columIndex]);
			
			//根据标题设置列宽
			int width = title.length() * pixel;
			sheet.setColumnWidth(columIndex, width < 5000 ? 5000 : width );
		}
    	// 遍历集合数据，产生数据行 第一行存放
		if(list != null) {
			for(int i = 0 ; i < list.size(); i++ ) {
	    		rowIndex++;
	    		HSSFRow contextRow = sheet.createRow(rowIndex);
	    		
	    		//遍历列 
	    		if(keys != null) {
	    			for(int j=0;j<keys.length;j++){
		    			
		    			Object val = list.get(i).get(keys[j]);
		    			HSSFCell contentCell = contextRow.createCell(j);
		    			
		    			Boolean isNum = false;//val 是否为数值型
		    			Boolean isInteger = false; //val 是否为整数
		    			Boolean isPercent = false; //val 是否为百分数
		    			
		    			if(val != null) {
		    				//判断data是否为数值型
		                    isNum = val.toString().matches("^(-?\\d+)(\\.\\d+)?$");
		                    //判断data是否为整数（小数部分是否为0）
		                    isInteger=val.toString().matches("^[-\\+]?[\\d]*$");
		                    //判断data是否为百分数（是否包含“%”）
		                    isPercent=val.toString().contains("%");
		    			} else {
		    				val = " ";
		    			}
		    			
		    			// 如果单元格内容是数值类型，涉及到金钱（金额、本、利），则设置cell的类型为数值型，设置data的类型为数值类型
		                if (isNum && !isPercent) {
		                    HSSFDataFormat df = workbook.createDataFormat(); // 此处设置数据格式
		                    if (isInteger) {
		                    	contextStyle.setDataFormat(df.getBuiltinFormat("#,#0"));//数据格式只显示整数
		                    	// 设置单元格格式
			                    contentCell.setCellStyle(contextStyle);
			                    // 设置单元格内容为double类型
			                    if(val.toString().length()<9){
			                    contentCell.setCellValue(Integer.parseInt(val.toString()));
			                    }else{
			                    	contentCell.setCellStyle(contextStyle);
				                    // 设置单元格内容为字符型
					                contentCell.setCellValue(val.toString());
			                    }
		                    } else {
		                    	contextStyle.setDataFormat(df.getBuiltinFormat("#,##0.00"));//保留两位小数点
		                    	// 设置单元格格式
			                    contentCell.setCellStyle(contextStyle);
			                    // 设置单元格内容为double类型
			                    contentCell.setCellValue(Double.parseDouble(val.toString()));
		                    }                   
		                } else {
		                    contentCell.setCellStyle(contextStyle);
		                    // 设置单元格内容为字符型
			                contentCell.setCellValue(val.toString());
		                }
		                // 设置单元格内容为字符型
		                // contentCell.setCellValue(val.toString());
		                
		                // 根据内容设置列宽度
		                int columWidth = val.toString().length() * pixel;
		                int w = sheet.getColumnWidth(j);
		                if(columWidth > w && columWidth < (20 * pixel)) {
		                	sheet.setColumnWidth(j , columWidth);
		                }
		    		}
	    		}
	    	}
		}
    	return workbook;
    }

/**
     * 导入Excel
     * 
     * @param clazz
     * @param xls
     * @return
     * @throws Exception
     */
    public static List importExcelXLS(Class<?> clazz, InputStream xls) throws Exception {
        try {
            // 取得Excel
            HSSFWorkbook wb = new HSSFWorkbook(xls);
            HSSFSheet sheet = wb.getSheetAt(0);
            Field[] fields = clazz.getDeclaredFields();
            List<Field> fieldList = new ArrayList<Field>(fields.length);
            for (Field field : fields) {
                if (field.isAnnotationPresent(ModelProp.class)) {
                    ModelProp modelProp = field.getAnnotation(ModelProp.class);
                    if (modelProp.colIndex() != -1) {
                        fieldList.add(field);
                    }
                }
            }
            // 行循环
            List<ImportModel> modelList = new ArrayList<ImportModel>(sheet.getPhysicalNumberOfRows() * 2);
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                // 数据模型
                ImportModel model = (ImportModel) clazz.newInstance();
                int nullCount = 0;
                Exception nullError = null;
                for (Field field : fieldList) {
                    ModelProp modelProp = field.getAnnotation(ModelProp.class);
                    HSSFCell cell = sheet.getRow(i).getCell(modelProp.colIndex());
                    try {
                        if (cell == null || cell.toString().length() == 0) {
                            nullCount++;
                            if (!modelProp.nullable()) {
                                nullError = new Exception(StringFormat.format(notnullerror,
                                        new String[] { "" + (1 + i), modelProp.name(), modelProp.name() }));

                            }
                        } else if (field.getType().equals(Date.class)) {
                            if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(), new Date(parseDate(parseStringXLS(cell))));
                            } else {
                                BeanUtils.setProperty(model, field.getName(),
                                        new Date(cell.getDateCellValue().getTime()));

                            }
                        } else if (field.getType().equals(Timestamp.class)) {
                            if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(),
                                        new Timestamp(parseDate(parseStringXLS(cell))));
                            } else {
                                BeanUtils.setProperty(model, field.getName(),
                                        new Timestamp(cell.getDateCellValue().getTime()));
                            }

                        } else if (field.getType().equals(java.sql.Date.class)) {
                            if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(),
                                        new java.sql.Date(parseDate(parseStringXLS(cell))));
                            } else {
                                BeanUtils.setProperty(model, field.getName(),
                                        new java.sql.Date(cell.getDateCellValue().getTime()));
                            }
                        } else if (field.getType().equals(Integer.class)) {
                            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(), (int) cell.getNumericCellValue());
                            } else if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(), Integer.parseInt(parseStringXLS(cell)));
                            }
                        } else if (field.getType().equals(BigDecimal.class)) {
                            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(),
                                        new BigDecimal(cell.getNumericCellValue()));
                            } else if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(), new BigDecimal(parseStringXLS(cell)));
                            }
                        } else {
                            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(),
                                        new BigDecimal(cell.getNumericCellValue()));
                            } else if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(), parseStringXLS(cell));
                            }
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                        throw new Exception(StringFormat.format(errormsg, new String[] { "" + (1 + i), modelProp.name() })
                                + "," + e.getMessage());
                    }
                }
                if (nullCount == fieldList.size()) {
                    break;
                }
                if (nullError != null) {
                    throw nullError;
                }
                modelList.add(model);
            }
            return modelList;

        } finally {
            xls.close();
        }
    }
    
    public static List importExcelXLSX(Class<?> clazz, InputStream xls) throws Exception {
        try {
            // 取得Excel
            XSSFWorkbook wb = new XSSFWorkbook(xls);
            XSSFSheet sheet = wb.getSheetAt(0);
            Field[] fields = clazz.getDeclaredFields();
            List<Field> fieldList = new ArrayList<Field>(fields.length);
            for (Field field : fields) {
                if (field.isAnnotationPresent(ModelProp.class)) {
                    ModelProp modelProp = field.getAnnotation(ModelProp.class);
                    if (modelProp.colIndex() != -1) {
                        fieldList.add(field);
                    }
                }
            }
            // 行循环
            List<ImportModel> modelList = new ArrayList<ImportModel>(sheet.getPhysicalNumberOfRows() * 2);
            for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
                // 数据模型
                ImportModel model = (ImportModel) clazz.newInstance();
                int nullCount = 0;
                Exception nullError = null;
                for (Field field : fieldList) {
                    ModelProp modelProp = field.getAnnotation(ModelProp.class);
                    XSSFCell cell = sheet.getRow(i).getCell(modelProp.colIndex());
                    try {
                        if (cell == null || cell.toString().length() == 0) {
                            nullCount++;
                            if (!modelProp.nullable()) {
                                nullError = new Exception(StringFormat.format(notnullerror,
                                        new String[] { "" + (1 + i), modelProp.name(), modelProp.name() }));
                            }
                        } else if (field.getType().equals(Date.class)) {
                            if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(), new Date(parseDate(parseStringXLSX(cell))));
                            } else {
                                BeanUtils.setProperty(model, field.getName(),
                                        new Date(cell.getDateCellValue().getTime()));
                            }
                        } else if (field.getType().equals(Timestamp.class)) {
                            if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(),
                                        new Timestamp(parseDate(parseStringXLSX(cell))));
                            } else {
                                BeanUtils.setProperty(model, field.getName(),
                                        new Timestamp(cell.getDateCellValue().getTime()));
                            }

                        } else if (field.getType().equals(java.sql.Date.class)) {
                            if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(),
                                        new java.sql.Date(parseDate(parseStringXLSX(cell))));
                            } else {
                                BeanUtils.setProperty(model, field.getName(),
                                        new java.sql.Date(cell.getDateCellValue().getTime()));
                            }
                        } else if (field.getType().equals(Integer.class)) {
                            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(), (int) cell.getNumericCellValue());
                            } else if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(), Integer.parseInt(parseStringXLSX(cell)));
                            }
                        } else if (field.getType().equals(BigDecimal.class)) {
                            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(),
                                        new BigDecimal(cell.getNumericCellValue()));
                            } else if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(), new BigDecimal(parseStringXLSX(cell)));
                            }
                        } else {
                            if (Cell.CELL_TYPE_NUMERIC == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(),
                                        new BigDecimal(cell.getNumericCellValue()));
                            } else if (Cell.CELL_TYPE_STRING == cell.getCellType()) {
                                BeanUtils.setProperty(model, field.getName(), parseStringXLSX(cell));
                            }
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                        throw new Exception(StringFormat.format(errormsg, new String[] { "" + (1 + i), modelProp.name() })
                                + "," + e.getMessage());
                    }
                }
                if (nullCount == fieldList.size()) {
                    break;
                }
                if (nullError != null) {
                    throw nullError;
                }
                modelList.add(model);
            }
            return modelList;

        } finally {
            xls.close();
        }
    }
    
    @Retention(RetentionPolicy.RUNTIME)
     @Target(ElementType.FIELD)
     public @interface ModelProp{
         public String name();
         public int colIndex() default -1;
         public boolean nullable() default true;
         public String interfaceXmlName() default "";
     }
    private static String parseStringXLS(HSSFCell cell) {
         return String.valueOf(cell).trim();
    }
    private static String parseStringXLSX(XSSFCell cell) {
        return String.valueOf(cell).trim();
   }
    private static long parseDate(String dateString) throws ParseException {
        if (dateString.indexOf("/") == 4) {
            return new SimpleDateFormat("yyyy/MM/dd").parse(dateString).getTime();
        } else if (dateString.indexOf("-") == 4) {
            return new SimpleDateFormat("yyyy-MM-dd").parse(dateString).getTime();
        } else if (dateString.indexOf("年") == 4) {
            return new SimpleDateFormat("yyyy年MM月dd").parse(dateString).getTime();
        } else if (dateString.length() == 8) {
            return new SimpleDateFormat("yyyyMMdd").parse(dateString).getTime();
        } else {
            return new Date().getTime();
        }
    }
    
    static class StringFormat {  
    	  
        public static String format(String str, String... args) {  
            for (int i = 0; i < args.length; i++) {  
                str = str.replaceFirst("\\{" + i + "\\}", args[i]);  
            }  
            return str;  
        }
    }

}