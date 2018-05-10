package com.songhj.controller;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import com.songhj.util.CommUtil;
import com.songhj.util.ExcelUtil;



/**
 * Excel导入导出操作
 * @author songhj
 *
 */
@RestController
public class IndexViewController {
	
	/**
	 * Excel数据导入
	 * @param request
	 * @param response
	 * @param filePro
	 * @return
	 */
	@RequestMapping(value="/importExcel")
	public Map<String, Object> importExcel(HttpServletRequest request,HttpServletResponse response, String filePro){
		Map<String, Object> map = new HashMap<>();
		String keys[] = {"name","age","sex"};
		try {
			List<Map<String,String>> listData = ExcelUtil.getExcelData(request, "file",keys);
			if(listData.size() == 0){
				map.put("status",-1);
				map.put("message","上传失败，上传数据必须大于一条");
				return map;
			}
			for (Map<String, String> dataMap : listData) {
				System.out.println(keys[0] + ":" + dataMap.get(keys[0]));
				System.out.println(keys[1] + ":" + dataMap.get(keys[1]));
				System.out.println(keys[2] + ":" + dataMap.get(keys[2]));
			}
			map.put("listData", listData);
			map.put("code", 1);
			map.put("message", "导入成功");
		} catch (Exception e) {
			e.printStackTrace();
		}
		return map;
	}

	/**
	 * 数据导出Excel
	 * @param request
	 * @param response
	 * @param columnNames
	 * @param keys
	 * @param methodName
	 * @return
	 * @throws IOException
	 */
	@RequestMapping(value = "/excelExport", method = { RequestMethod.GET, RequestMethod.POST })
	public Map<String,Object> excelExport(HttpServletRequest request, HttpServletResponse response) throws IOException {
		Map<String,Object> data = new HashMap<>();
		String fileName = CommUtil.formatTime("yyyyMMddHHmmss", new Date()) +".xls";
		
		String columnNames[] = {"姓名","年龄","性别"};
		String keys[] = {"name","age","sex"};
		
		List<Map<String,Object>> listMap = new ArrayList<>();
		Map<String,Object> map = new HashMap<>();
		map.put("name", "jack");
		map.put("age", "18");
		map.put("sex", "男");
		listMap.add(map);
		map = new HashMap<>();
		map.put("name", "tom");
		map.put("age", "20");
		map.put("sex", "男");
		listMap.add(map);
		try {
			//创建Workbook
			Workbook wb = ExcelUtil.createWorkBook(listMap, keys, columnNames);
			//保存路径
			String savePath = request.getServletContext().getRealPath("/") + File.separator + fileName;
			// 创建文件流
			OutputStream stream = new FileOutputStream(savePath);
			// 写入数据
			wb.write(stream);
			// 关闭文件流
			stream.close();
			
			//返回结果
			data.put("code", 1);
			String downloadUrl = request.getScheme() + "://"+request.getServerName() + ":" + 
					request.getServerPort() + "/" + fileName;
			data.put("download", downloadUrl);
			data.put("message", "文件流输出成功");
			
			System.out.println("\n数据导出成功，下载路劲：" + downloadUrl);
		} catch (IOException e) {
			System.err.println(e.getMessage());
			data.put("code", -1);
			data.put("message", "下载出错");
			return data;
		}
		return data;
	}
}
