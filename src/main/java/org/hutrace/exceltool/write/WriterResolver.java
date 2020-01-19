package org.hutrace.exceltool.write;

import org.apache.poi.ss.usermodel.Row;

/**
 * <p>写入Excel数据的数据解析器标准接口
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 */
public interface WriterResolver {
	
	/**
	 * 写入标题数据
	 * @param row {@link Row}行对象
	 */
	void title(Row row);
	
	/**
	 * 写入非标题到的行数据
	 * @param row {@link Row}行对象
	 * @param index 行下标
	 */
	void row(Row row, int index);
	
}
