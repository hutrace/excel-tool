package org.hutrace.exceltool.utils;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hutrace.exceltool.pojo.ExcelType;

/**
 * <p>Excel工具包的常用方法公共类
 * <p>包含一些在读取与写入可以使用的公共方法
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月14日
 */
public class ExcelCommon {
	
	private static final String XLS = "xls";
	private static final String XLSX = "xlsx";
	
	/**
	 * 根据文件绝对路径获取Excel的文件类型
	 * @param path 绝对路径获取
	 * @return Excel的文件类型
	 */
	public ExcelType getType(String path) {
		String sux = path.substring(path.lastIndexOf(".") + 1);
		if(XLS.equalsIgnoreCase(sux)) {
			return ExcelType.XLS;
		}else if(XLSX.equalsIgnoreCase(sux)) {
			return ExcelType.XLSX;
		}else {
			return null;
		}
	}
	
	/**
	 * <p>使用in创建{@link Workbook}对象，此方法用于读取时调用
	 * <p>当type为null时会抛出{@link NullPointerException}异常
	 * @param in 输入流
	 * @param type Excel文件类型
	 * @return {@link Workbook}
	 * @throws IOException
	 */
	public Workbook createWorkbook(InputStream in, ExcelType type) throws IOException {
		if(type == null) {
			throw new NullPointerException("The [type] cannot be null");
		}
		if(type == ExcelType.XLS) {
			return new HSSFWorkbook(in);
		}else {
			return new XSSFWorkbook(in);
		}
	}
	
	/**
	 * <p>创建{@link Workbook}对象，此方法用于写入时调用
	 * <p>当type为null时会抛出{@link NullPointerException}异常
	 * @param in 输入流
	 * @param type Excel文件类型
	 * @return {@link Workbook}
	 * @throws IOException
	 */
	public Workbook createWorkbook(ExcelType type) {
		if(type == null) {
			throw new NullPointerException("The [type] cannot be null");
		}
		if(type == ExcelType.XLS) {
			return new HSSFWorkbook();
		}else {
			return new XSSFWorkbook();
		}
	}
	
}
