package com.songhj.util;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

/**
 * Excel导出
 * @author song
 *
 */
public class ExcelUtil {
	/**
	 *
	 * @param list 导出的数据
	 * @param keys 取每列数据的key
	 * @param columnNames 列名
	 * @return
	 */
    public static Workbook createWorkBook(List<Map<String, Object>> listMap, String keys[], String columnNames[]) {
        // 创建excel工作簿
        Workbook wb = new HSSFWorkbook();
        // 创建第一个sheet（页），并命名
        Sheet sheet = wb.createSheet("Sheet1");
        // 手动设置列宽。第一个参数表示要为第几列设；，第二个参数表示列的宽度，n为列高的像素数。
        for(int i = 0; i < keys.length; i++){
            sheet.setColumnWidth((short) i, (short) (35 * 100));
        }

   /**
        // 创建两种单元格格式
        CellStyle cs = wb.createCellStyle();
        CellStyle cs2 = wb.createCellStyle();

        // 创建两种字体
        Font f = wb.createFont();
        Font f2 = wb.createFont();

        // 创建第一种字体样式（用于列名）
        f.setFontHeightInPoints((short) 10);
        f.setColor(IndexedColors.BLACK.getIndex());
        f.setBoldweight(Font.BOLDWEIGHT_BOLD);

        // 创建第二种字体样式（用于值）
        f2.setFontHeightInPoints((short) 10);
        f2.setColor(IndexedColors.BLACK.getIndex());

        // 设置第一种单元格的样式（用于列名）
        cs.setFont(f);
        cs.setBorderLeft(CellStyle.BORDER_THIN);
        cs.setBorderRight(CellStyle.BORDER_THIN);
        cs.setBorderTop(CellStyle.BORDER_THIN);
        cs.setBorderBottom(CellStyle.BORDER_THIN);
        cs.setAlignment(CellStyle.ALIGN_CENTER);

        // 设置第二种单元格的样式（用于值）
        cs2.setFont(f2);
        cs2.setBorderLeft(CellStyle.BORDER_THIN);
        cs2.setBorderRight(CellStyle.BORDER_THIN);
        cs2.setBorderTop(CellStyle.BORDER_THIN);
        cs2.setBorderBottom(CellStyle.BORDER_THIN);
        cs2.setAlignment(CellStyle.ALIGN_CENTER);

   */
        // 创建第一行
        Row row = sheet.createRow((short) 0);

        //设置列名
        for(int i = 0; i < columnNames.length; i++){
            Cell cell = row.createCell(i);
            cell.setCellValue(columnNames[i]);
            /**cell.setCellStyle(cs);*/
        }

        //设置每行每列的值
        for (short i = 0; i < listMap.size(); i++) {
            // Row 行,Cell 方格 , Row 和 Cell 都是从0开始计数的
            // 创建一行，在页sheet上
            Row row1 = sheet.createRow((short) i + 1);
            // 在row行上创建一个方格
            for(short j = 0; j < keys.length; j++){
                Cell cell = row1.createCell(j);
                cell.setCellValue(listMap.get(i).get(keys[j]) == null?" " : listMap.get(i).get(keys[j]).toString());
                /**cell.setCellStyle(cs2);*/
            }
        }

        return wb;
    }



    /**
     * 获取导入Excel的数据
     * @param request
     * @param filePro
     * @param keys
     * @return
     */
    public static List<Map<String,String>> getExcelData(HttpServletRequest request,String filePro,String keys[]){
    	List<Map<String,String>> data = new ArrayList<>();
    	try {
    		if(!CommUtil.isNotNull(filePro)){
    			filePro = "file";
    		}
    		MultipartHttpServletRequest multipartRequest = (MultipartHttpServletRequest) request;
    	    MultipartFile file = (MultipartFile) multipartRequest.getFile(filePro);
    	    //获得文件名
            String fileName = file.getOriginalFilename();
			Workbook wb = null;
			if(fileName.endsWith("xlsx")){
				wb = new XSSFWorkbook(file.getInputStream());
			}else{
				wb = new HSSFWorkbook(file.getInputStream());
			}
			//获得第一个表单
			Sheet sheet = wb.getSheetAt(0);
			//获得第一个表单的迭代器
            Iterator<Row> rows = sheet.rowIterator();
            int i = 0;
            while (rows.hasNext()) {
            	i ++;
            	Map<String,String> rowMap = new HashMap<>();
            	//获得行数据
                Row row = rows.next();
                //跳过头部
                if(row.getRowNum() == 0){
                	continue;
                }
                Iterator<Cell> cells = row.cellIterator();
                //获得行的迭代器
                int j = 0,k = 0;
                while (cells.hasNext()) {
                	Cell cell = cells.next();
                	//类型判断
                	String key = "";
                	//防止越界
                	if(keys.length > cell.getColumnIndex()){
                		key = keys[cell.getColumnIndex()];
                	}
                	if(CommUtil.isNotNull(key)){
                		String value = formatCell(cell);
                		rowMap.put(key, value);
                		if(!CommUtil.isNotNull(value)){
                			j ++; //记录空值得数量
                		}
                		k ++; //记录多少列
                	}
                }
                //如果i=j，说明一行都是空的
                if(j == k){
                	break;
                }else{
                	data.add(rowMap);
                }
                if(i > 50001){
                	System.out.println("\n============导入数据大于五万条，立即停止===============");
                	break;
                }
            }
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    	return data;
    }

    /**
     * 反射机制 map转实体
     * @param classPath java类路径
     * @param listData  map数据
     * @return
     */
    public static List<Object> factoryMapToJavaObj(String classPath,List<Map<String, String>> listData){
    	List<Object> list = new ArrayList<Object>();
    	try {
    		//获取类
    		Class<?> c = Class.forName(classPath);
    		//获取类的所有属性
    		Field[] fs = c.getDeclaredFields();
    		Object valObj = null;
    		for (Map<String, String> mapData : listData)
    		{
    			//获取类的一个实例
    			Object o = c.newInstance();
    			for (Field field : fs)
    			{
    				String fieldName = field.getName();
    				String fieldType = field.getType().getName();
    				String value = mapData.get(fieldName);
    				if(CommUtil.isNotNull(value)){
    					valObj = value.trim();
    				}else{
    					valObj = null;
    				}
    				if(Modifier.toString(field.getModifiers()).indexOf("static")!=-1||fieldName.equals("id")){continue;}
    				if(fieldType.equals("java.math.BigDecimal")){
    					if(CommUtil.isNotNull(valObj)){
    						if(valObj.toString().indexOf("%")!=-1){
    							String newVal = valObj.toString().replaceAll("%", "");
    							Double num = Double.valueOf(newVal)/100;
    							valObj = BigDecimal.valueOf(num);
    						}
    						else{
    							valObj = BigDecimal.valueOf(Double.valueOf(value.replace(",", "")));
    						}
    					}else{
    						valObj = BigDecimal.valueOf(0);
    					}
    				}
    				if(fieldType.equals("java.util.Date")&&CommUtil.isNotNull(value)){
    					valObj = CommUtil.formatDate(value.replace("/", "-") ,"yyyy-MM-dd");
    				}
    				if(fieldType.equals("java.lang.Integer")){
    					if(CommUtil.isNotNull(valObj)){
    						valObj = Integer.valueOf(valObj.toString());
    					}else{
    						valObj = Integer.valueOf(0);
    					}
    				}
    				if(fieldType.equals("java.lang.Long")){
    					if(CommUtil.isNotNull(valObj)){
    						valObj = Long.valueOf(valObj.toString());
    					}else{
    						valObj = Long.valueOf(0);
    					}
    				}
    				if(CommUtil.isNotNull(valObj)){
    					//设置可访问私有属性
    					field.setAccessible(true);
    					//给o对象的属性赋值
    					field.set(o, valObj);
    				}
    			}
    			list.add(o);
    		}
    	} catch (ClassNotFoundException | InstantiationException | IllegalAccessException e) {
    		e.printStackTrace();
    	}
    	return list;
    }

    public static void main(String[] args) {
//    	Map<String, String> map = new HashMap<>();
//    	List l = new ArrayList<>();
//    	l.add(map);
//    	new ExcelUtil().factoryMapToJavaObj("com.qhiex.foundation.domain.po.system.BusMediInterface", l);
	}


    /**
     * 按类型取值
     * @param cell
     * @return
     */
    public static String formatCell(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {

             //数值格式
        case HSSFCell.CELL_TYPE_NUMERIC:
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
            	//如果是日期格式
                return CommUtil.formatShortDate(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
            }else{
            	//字符时
            	cell.setCellType(HSSFCell.CELL_TYPE_STRING);
            	return cell.getStringCellValue();
            }

            //字符串
        case HSSFCell.CELL_TYPE_STRING:
            return cell.getStringCellValue();

            // 公式
        case HSSFCell.CELL_TYPE_FORMULA:
            return cell.getCellFormula();

            // 空白
        case HSSFCell.CELL_TYPE_BLANK:
            return "";

            // 布尔取值
        case HSSFCell.CELL_TYPE_BOOLEAN:
            return cell.getBooleanCellValue() + "";

            //错误类型
        case HSSFCell.CELL_TYPE_ERROR:
            break;
        }
        return "";
    }
}
