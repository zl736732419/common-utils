package com.zheng.test;

import com.zheng.utils.BarCodeUtil;
import com.zheng.utils.QrCodeUtil;

import java.io.File;
import java.io.IOException;

/**
 * 二维码、一维码测试
 * Created by zhenglian on 2016/12/30.
 */
public class CodeTest {
	public static void main(String[] args) throws Exception {
//		new CodeTest().createBarCode();
//        new CodeTest().createQR();
        new CodeTest().parseQR();
	}
	
	private void parseQR() throws IOException {
		File file = new File("D://images/zxing.png");
		QrCodeUtil.parseQR(file);
	}
	
	private void createQR() throws Exception {
		String msg = "123456789";
		String name = "小张";
		String path = "D://images/zxing.png";
		QrCodeUtil.createQR(msg, name, path);
	}
	
	private void createBarCode() throws Exception {
		String msg = "1234567890";
		String name = "老王";
		
		String path = "D:/images/output.png";
		BarCodeUtil.create(msg, name, path);
	}
	
}
