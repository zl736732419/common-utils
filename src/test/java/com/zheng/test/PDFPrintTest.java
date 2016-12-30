package com.zheng.test;

import com.google.common.collect.Maps;
import com.google.common.io.Resources;
import com.zheng.utils.PDFPrintUtil;

import java.io.File;
import java.net.URI;
import java.util.Map;

/**
 * Created by zhenglian on 2016/12/30.
 */
public class PDFPrintTest {
	public static void main(String[] args) throws Exception {
		new PDFPrintTest().addInfoToPDF();
//        new PDFPrintTest().mergePdfs();
		
	}
	
	private void addInfoToPDF() throws Exception {
		String outFile = "D:/test.pdf";
		URI uri = Resources.getResource("template.pdf").toURI();
		File templateFile = new File(uri);
		Map<String, Object> params = Maps.newHashMap();
		params.put("username", "小赵");
		params.put("school", "南山实验中学");
		
		String imagePath = "d://images/zxing.png";
		PDFPrintUtil.createPdfByTemplate(templateFile.getAbsolutePath(), outFile, params, imagePath);
	}
	
	private void mergePdfs() throws Exception {
		//1000个学生为一组
//        String dirPath = "C:/Users/dell/Desktop/pdfs-A4";
//        String outPath = "C:/Users/dell/Desktop/merge-A4.pdf";
		
		//1000个学生为一组
		URI uri = Resources.getResource("pdfs").toURI();
		File dir = new File(uri);
		String outPath = "D:/merge.pdf";
		
		
		PDFPrintUtil.mergePdfs(dir.getAbsolutePath(), outPath);
	}
}
