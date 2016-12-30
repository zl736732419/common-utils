package com.zheng.utils;

import com.google.common.base.Preconditions;
import com.google.common.base.Strings;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;


/**
 * Excel表格工具类
 * 采用apache poi 工具操作
 * 流程：
 * createWorkbook().createSheet().createRow().createCell().mergeCells().writeExcelToDisk()
 */
public class ExcelUtil {
	
	protected static DecimalFormat decimalFormat = new DecimalFormat("#.######");
	protected static SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	
	private CellStyle cs;
	
	/**
	 * 生成创建xls格式的excel工作簿
	 *
	 * @return
	 * @throws Exception
	 */
	public Workbook createWorkbook() throws Exception {
		Workbook wb = new HSSFWorkbook();
		cs = null;
		this.getCellStyle(wb); //每一次创建工作簿就重新更新单元格样式
		return wb;
	}
	
	/**
	 * 创建excel工作页
	 *
	 * @param wb
	 * @param sheetName
	 * @return
	 */
	public Sheet createSheet(Workbook wb, String sheetName) {
		Preconditions.checkNotNull(wb, "工作簿不能为空");
		Preconditions.checkNotNull(sheetName, "工作簿名称不能为空");
		
		Sheet sheet = wb.createSheet(sheetName);
		return sheet;
	}
	
	/**
	 * 获取excel行
	 *
	 * @param sheet
	 * @param rowNum
	 * @return
	 * @author zhenglian
	 * @data 2016年12月22日 下午2:20:00
	 */
	public org.apache.poi.ss.usermodel.Row getRow(Sheet sheet, int rowNum) {
		Preconditions.checkNotNull(sheet, "工作页面不能为空");
		Preconditions.checkArgument(rowNum >= 0, "行号必须大于0");
		org.apache.poi.ss.usermodel.Row row = sheet.getRow(rowNum);
		if (null == row) {
			row = sheet.createRow(rowNum);
			row.setHeightInPoints(12.75f);
		}
		
		return row;
	}
	
	/**
	 * 获取单元格样式,单元格样式数量有限制，这里设置成单例共享同一个样式
	 *
	 * @return
	 * @author zhenglian
	 * @data 2016年12月22日 下午2:09:30
	 */
	private CellStyle getCellStyle(Workbook wb) {
		if (cs == null) {
			cs = wb.createCellStyle();
			cs.setAlignment(HorizontalAlignment.CENTER);
			cs.setVerticalAlignment(VerticalAlignment.CENTER);
			
			Font font = wb.createFont();
			font.setFontName("宋体");
			font.setItalic(false);
			font.setBold(false);
			font.setColor(Font.COLOR_NORMAL);
			font.setFontHeightInPoints((short) 9);
			font.setStrikeout(false);
			cs.setFont(font);
			
			cs.setBorderTop(BorderStyle.THIN);
			cs.setBorderRight(BorderStyle.THIN);
			cs.setBorderBottom(BorderStyle.THIN);
			cs.setBorderLeft(BorderStyle.THIN);
		}
		
		return cs;
	}
	
	/**
	 * 生成表格中的一列
	 *
	 * @param wb     工作簿
	 * @param row    行
	 * @param column 列
	 * @param text   列内容
	 * @return
	 * @throws Exception
	 */
	public Cell createCell(Workbook wb, org.apache.poi.ss.usermodel.Row row, int column, Object text) throws Exception {
		Preconditions.checkNotNull(wb, "工作簿不能为空");
		Preconditions.checkNotNull(row, "工作行不能为空");
		Preconditions.checkArgument(column >= 0, "列号必须大于0");
		
		Cell cell = row.createCell(column);
		setCellValue(cell, text);
		CellStyle cs = getCellStyle(wb);
		cell.setCellStyle(cs);
		
		return cell;
	}
	
	/**
	 * 根据类型设置单元格值
	 *
	 * @param cell
	 * @param text
	 * @author zhenglian
	 * @data 2016年12月23日 上午10:49:03
	 */
	private void setCellValue(Cell cell, Object text) {
		Preconditions.checkNotNull(cell, "单元格不能为空");
		
		if (null == text) {
			text = "";
		}
		
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		if (text instanceof Date) {
			cell.setCellValue(df.format((Date) text));
		} else if (text instanceof Boolean) {
			cell.setCellValue((Boolean) text);
		} else if (text instanceof Calendar) {
			Calendar c = (Calendar) text;
			String str = getDateTime(c);
			cell.setCellValue(str);
		} else if (text instanceof Double) {
			cell.setCellValue((Double) text);
		} else {
			String str = text.toString();
			if (isNumber(str)) {
				cell.setCellValue(Double.parseDouble(str));
			} else {
				cell.setCellValue(str);
			}
		}
	}
	
	/**
	 * 根据日历获取标准格式时间 年-月-日 时:分:秒
	 *
	 * @param c
	 * @return
	 * @author zhenglian
	 * @data 2016年12月23日 上午11:03:37
	 */
	private String getDateTime(Calendar c) {
		int year = c.get(Calendar.YEAR);
		int month = c.get(Calendar.MONTH) + 1;// 月
		int day = c.get(Calendar.DATE); // 日
		
		int hours = c.get(Calendar.HOUR);
		int minutes = c.get(Calendar.MINUTE);
		int seconds = c.get(Calendar.SECOND);
		
		String str = new StringBuilder().append(year).append("-").append(formatTime(month)).append("-")
				.append(formatTime(day)).append(" ").append(formatTime(hours)).append(":").append(formatTime(minutes))
				.append(":").append(formatTime(seconds)).toString();
		return str;
	}
	
	/**
	 * 标准化时间 小于10的需要加上0 比如 九分8秒 -> 09:08
	 *
	 * @param time
	 * @return
	 * @author zhenglian
	 * @data 2016年12月23日 上午11:04:32
	 */
	private String formatTime(int time) {
		String str = time + "";
		if (time < 10) {
			str = "0" + str;
		}
		
		return str;
	}
	
	/**
	 * 判断字符串是否是数字格式
	 *
	 * @param text
	 * @return
	 * @author zhenglian
	 * @data 2016年12月23日 上午10:50:48
	 */
	private boolean isNumber(String text) {
		if (Strings.isNullOrEmpty(text)) {
			return false;
		}
		
		boolean isNumber = true;
		try {
			Double.parseDouble(text);
		} catch (NumberFormatException e) {
			isNumber = false;
		}
		
		return isNumber;
	}
	
	/**
	 * 生成表格中的一列
	 *
	 * @param wb     工作簿
	 * @param row    行
	 * @param column 列
	 * @param texts   多列内容
	 * @return
	 * @throws Exception
	 */
	public void createCells(Workbook wb, org.apache.poi.ss.usermodel.Row row, int column, Object... texts)
			throws Exception {
		Preconditions.checkNotNull(wb, "工作簿不能为空");
		Preconditions.checkNotNull(row, "工作行不能为空");
		Preconditions.checkNotNull(texts, "行不能为空");
		for (int i = column, idx = 0, size = texts.length + column; i < size; i++) {
			createCell(wb, row, i, texts[idx++]);
		}
	}
	
	/**
	 * 从第0列开始添加多列
	 * @param wb
	 * @param row
	 * @param texts
	 * @throws Exception
	 */
	public void createCells(Workbook wb, org.apache.poi.ss.usermodel.Row row, Object... texts) throws Exception {
		createCells(wb, row, 0, texts);
	}
	
	/**
	 * 合并单元格 需要注意的一点是所合并的单元格必须要先创建后再合并 比如要合并(1,1),(2,1) 则要先创建这两列，不然出问题
	 *
	 * @param sheet
	 * @param firstRowIdx
	 * @param lastRowIdx
	 * @param firstColIdx
	 * @param lastColIdx
	 * @param bolder
	 * @param text
	 */
	public void mergeCells(Sheet sheet, int firstRowIdx, int lastRowIdx, int firstColIdx, int lastColIdx,
	                       boolean bolder, String text) throws Exception {
		Preconditions.checkNotNull(sheet, "工作页不能为空");
		Preconditions.checkArgument(firstRowIdx >= 0 && lastRowIdx >= 0 && firstColIdx >= 0 && lastColIdx >= 0,
				"行列号必须大于0");
		
		if (null == text) {
			text = "";
		}
		
		org.apache.poi.ss.usermodel.Row firstRow = sheet.getRow(firstRowIdx);
		firstRow = this.getRow(sheet, firstRowIdx);
		org.apache.poi.ss.usermodel.Row lastRow = sheet.getRow(lastRowIdx);
		lastRow = this.getRow(sheet, lastRowIdx);
		Cell cell;
		for (int i = firstColIdx; i <= lastColIdx; i++) {
			cell = firstRow.getCell(i, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
			if (null == cell) {
				cell = this.createCell(sheet.getWorkbook(), firstRow, i, text);
			}
			if (i == firstColIdx) {
				cell.setCellValue(text);
			}
		}
		
		for (int i = firstRowIdx; i <= lastColIdx; i++) {
			cell = lastRow.getCell(i, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_NULL_AND_BLANK);
			if (null == cell) {
				this.createCell(sheet.getWorkbook(), lastRow, i, text);
			}
		}
		
		sheet.addMergedRegion(new CellRangeAddress(firstRowIdx, lastRowIdx, firstColIdx, lastColIdx));
	}
	
	/**
	 * 将Excel表格导出到指定位置
	 *
	 * @param wb
	 * @param path
	 * @throws Exception
	 */
	public void writeExcelToDisk(Workbook wb, String path) throws Exception {
		Preconditions.checkNotNull(wb, "工作簿不能为空");
		Preconditions.checkNotNull(path, "Excel生成路径不能为空");
		
		FileOutputStream output = new FileOutputStream(path);
		wb.write(output);
		wb.close();
		output.close();
	}
	
	/**
	 * 获取合并的单元格的合并列范围
	 *
	 * @param sheet
	 * @param cell
	 * @return
	 */
	public CellRangeAddress getMergedRegion(Sheet sheet, Cell cell) {
		Preconditions.checkNotNull(sheet, "工作页不能为空");
		Preconditions.checkNotNull(cell, "列不能为空");
		
		int row = cell.getRowIndex();
		int column = cell.getColumnIndex();
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return range;
				}
			}
		}
		return null;
	}
	
	/**
	 * 解析单元格中的内容
	 * @param cell
	 * @return
	 */
	public String parse(Cell cell) {
		if (cell == null)
			return "";
		String result = new String();
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_NUMERIC:// 数字类型
				if (DateUtil.isCellDateFormatted(cell)) {// 处理日期格式、 时间格式
					Date date = cell.getDateCellValue();
					result = dateFormat.format(date);
				} else {
					double va = cell.getNumericCellValue();
					if (va == (int) va)// 去掉数值类型后面的".0"
						result = String.valueOf((int) va);
					else
						// result = String.valueOf(va); //if the double value is too
						// big, it will be displayed in E-notation
						result = decimalFormat.format(va);
				}
				break;
			case Cell.CELL_TYPE_FORMULA:
				// cell.getCellFormula();
				try {
					result = String.valueOf(cell.getNumericCellValue());
				} catch (IllegalStateException e) {
					result = String.valueOf(cell.getRichStringCellValue()).trim();
				}
				break;
			case Cell.CELL_TYPE_STRING:// String类型
				result = cell.getRichStringCellValue().toString().trim();
				break;
			case Cell.CELL_TYPE_BLANK:
				result = "";
				break;
			default:
				result = "";
				break;
		}
		
		return result.trim();
	}
	
	
}
