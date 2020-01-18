package org.hutrace.exceltool.read;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.hutrace.exceltool.exception.ExcelReaderException;

/**
 * <p>读取数据解析器的抽象类
 * <p>实现{@link ReaderResolver}接口，但不实现{@link #title(org.apache.poi.ss.usermodel.Row)}和{@link #row(org.apache.poi.ss.usermodel.Row, int)}方法。
 * <p>定义{@link #cell(Cell, int, int)}方法用于解析{@link Cell}的数据
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月14日
 * @param <T>
 * @see ReaderResolver
 */
public abstract class AbstractReaderResolver<T> implements ReaderResolver {
	
	public AbstractReaderResolver() {
		list = new ArrayList<>();
	}
	
	/**
	 * 数据集合，最终读取的数据。
	 */
	private List<T> list;
	
	@SuppressWarnings("unchecked")
	@Override
	public List<T> data() {
		return list;
	}
	
	/**
	 * 集合里面添加数据对象
	 * @param obj
	 */
	public void add(T obj) {
		list.add(obj);
	}

	/**
	 * <p>获取cell的数据
	 * <p>debug方法
	 * <p>通过{@link Cell#getCellType()}来判断取值类型
	 * @param cell 列对象
	 * @param rowIndex 列对象
	 * @param cellIndex 列对象
	 * @return 一行中某一列的数据
	 */
	public Object cell(Cell cell, int rowIndex, int cellIndex) {
		if(cell == null) {
			return "";
		}
		try {
			return cell(cell);
		}catch (Exception e) {
			throw new ExcelReaderException("Read data failed, row " + rowIndex + ", column " + cellIndex, e);
		}
	}
	
	/**
	 * <p>获取cell的数据
	 * <p>通过{@link Cell#getCellType()}来判断取值类型
	 * @param cell 列对象
	 * @return 一行中某一列的数据
	 */
	private Object cell(Cell cell) {
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				return "";
			case Cell.CELL_TYPE_NUMERIC:
				if(DateUtil.isCellDateFormatted(cell)) {
					return cell.getDateCellValue();
				}else {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					return cell.getStringCellValue();
				}
			case Cell.CELL_TYPE_STRING:
				return cell.getStringCellValue();
			case Cell.CELL_TYPE_FORMULA:
				return cell.getCellFormula();
			case Cell.CELL_TYPE_BOOLEAN:
				return cell.getBooleanCellValue();
			case Cell.CELL_TYPE_ERROR:
			default:
				return "";
		}
	}

}
