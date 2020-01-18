package org.hutrace.exceltool.read;

import java.util.List;

import org.apache.poi.ss.usermodel.Row;

/**
 * <p>读取Excel数据的数据解析器标准接口
 * <p>可以通过实现此接口来创建对应的数据解析器
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @param <T>
 * @time 2020年1月14日
 */
public interface ReaderResolver {
	
	/**
	 * 读取到整行title数据时触发的方法
	 * @param row
	 */
	void title(Row row);
	
	/**
	 * 读取到整行数据时触发的方法，title不会再此触发。
	 * @param row
	 */
	void row(Row row, int index);
	
	/**
	 * 获取读取的数据集合
	 * @return
	 */
	<T> List<T> data();
	
}
