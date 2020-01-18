package org.hutrace.exceltool.write;

import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.hutrace.exceltool.exception.ExcelWriteException;
import org.hutrace.exceltool.utils.TypeUtils;

/**
 * <p>写入Excel数据的数据解析器抽象类
 * <p>定义初始化数据方法、{@link #getData(int)}方法以及{@link #cell(Cell, Object, int, int)}方法
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月15日
 * @param <T>
 */
public abstract class AbstractWriterResolver<T> implements WriterResolver {
	
	/**
	 * 需要写入Excel的数据
	 * @param data
	 */
	private List<T> data;
	
	/**
	 * <p>创建对象，初始化需要写入的数据
	 * @param data 需要写入Excel的数据
	 */
	public AbstractWriterResolver(List<T> data) {
		this.data = data;
	}
	
	/**
	 * 根据下标获取一条数据
	 * @param i 需要获取数据的下标
	 * @return 一条数据
	 */
	protected T get(int i) {
		return data.get(i);
	}
	
	/**
	 * <p>向Excel的某一行中的某一列（就是一个格子）写入数据
	 * <p>此方法为debug方法，判断空与捕获写入时异常
	 * <p>详细过程见{@link #cell(Cell, Object)}
	 * @param cell 某一行中的某一列（格子）对象
	 * @param value 向格子中写入的数据
	 * @param rowIndex 行的下标
	 * @param cellIndex 列的下标
	 * @see #cell(Cell, Object)
	 */
	public void cell(Cell cell, Object value, int rowIndex, int cellIndex) {
		if(value == null) {
			cell.setCellValue("");
		}else {
			try {
				cell(cell, value);
			}catch (Exception e) {
				throw new ExcelWriteException("Error writing data, row " + rowIndex + ", column " + cellIndex, e);
			}
		}
	}
	
	/**
	 * <p>向Excel的某一行中的某一列（就是一个格子）写入数据
	 * <p>写入共分为四种类型，根据value的类型判断
	 * <pre>
	 *  1.double：所有数字类型都将转为double
	 *  2.boolean：对应布尔类型
	 *  3.Date：对应java.util.Date类型
	 *  4.String：剩下的均为String类型
	 * @param cell 某一行中的某一列（格子）对象
	 * @param value 向格子中写入的数据
	 */
	private void cell(Cell cell, Object value) {
		Class<?> clazs = value.getClass();
		if(clazs == int.class || clazs == Integer.class
				|| clazs == long.class || clazs == Long.class
				|| clazs == short.class || clazs == Short.class
				|| clazs == byte.class || clazs == Byte.class
				|| clazs == float.class || clazs == Float.class
				|| clazs == double.class || clazs == Double.class) {
			cell.setCellValue(TypeUtils.castToDouble(value));
		}else if(clazs == boolean.class || clazs == Boolean.class) {
			cell.setCellValue(TypeUtils.castToBoolean(value));
		}else if(clazs == Date.class) {
			cell.setCellValue((Date) value);
		}else {
			cell.setCellValue(value.toString());
		}
	}
	
}
