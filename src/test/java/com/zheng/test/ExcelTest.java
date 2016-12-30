package com.zheng.test;

import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import com.google.common.collect.Lists;
import com.zheng.utils.ExcelUtil;

/**
 * Excel测试
 * Created by zhenglian on 2016/12/19.
 */
public class ExcelTest {
	
	private ExcelUtil excelUtil = new ExcelUtil();
	
	@Test
	public void test() throws Exception {
		Workbook wb = excelUtil.createWorkbook();
		Sheet sheet = excelUtil.createSheet(wb, "高二美术");
		Row row1 = excelUtil.getRow(sheet, 0);
		Row row2 = excelUtil.getRow(sheet, 1);
		List<String> headers = Lists.newArrayList("考号", "姓名", "学校名称", "班级名称", "文理分科");
		for (int i = 0; i < headers.size(); i++) {
			excelUtil.mergeCells(sheet, row1.getRowNum(), row2.getRowNum(), i, i, false, headers.get(i));
		}
		
		List<String> headers2 = Lists.newArrayList("语文(文)", "文数", "英语(文)", "政治", "历史", "地理", "政史地", "语数外", "总分");
		List<String> secondHeaders = Lists.newArrayList("成绩", "市排名", "区排名", "校排名", "班排名");
		
		int start = headers.size();
		for (int i = 0; i < headers2.size(); i++) {
			excelUtil.mergeCells(sheet, row1.getRowNum(), row1.getRowNum(), start + i * secondHeaders.size(), 
					start + (i+1) * secondHeaders.size() - 1, false, headers2.get(i));
			for(int j = 0; j < secondHeaders.size(); j++) {
				excelUtil.createCell(wb, row2, start + secondHeaders.size() * i + j, secondHeaders.get(j));	
			}
		}
		
		String path = "D:/test.xls";
		excelUtil.writeExcelToDisk(wb, path);
		
		
	}
}
