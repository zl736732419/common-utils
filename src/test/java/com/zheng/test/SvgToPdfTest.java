package com.zheng.test;

import com.google.common.io.Resources;
import com.zheng.utils.SVGConverterUtil;

import java.io.File;
import java.net.URI;

/**
 * svg转pdf测试
 * Created by zhenglian on 2016/12/30.
 */
public class SvgToPdfTest {
	
	public static void main(String[] args) throws Exception {
		
		URI uri = Resources.getResource("test.svg").toURI();
		File svgFile = new File(uri);
		
		String name = svgFile.getName();
		name = name.substring(0, name.lastIndexOf("."));
		SVGConverterUtil converter = new SVGConverterUtil();
		converter.svg2PDF(svgFile, new File("D:/" + name + "_SVG文件转PDF.pdf"));
		
		// converter.svg2PNG(svgFile, new File("D:/"+name+"_SVG文件转PNG.PNG"));
		// converter.svg2JPEG(svgFile, new File("D:/" + name +
		// "_SVG文件转JPG.jpg"));
		
		// String svgCode = converter.svg2String(new File(svgpath));
		// converter.svg2PDF(svgCode, "D:/" + name + "_SVG代码转PDF.pdf");
		// converter.svg2PDF(svgCode, new FileOutputStream(new File("D:/" + name
		// + "_SVG代码转输出流.pdf")));
	}
	
}
