package org.hutrace.exceltool.write;

import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;

/**
 * <p>Excel写入数据解析器
 * <p>Map有序数据解析器
 * <p>此解析器写入Excel的列是有序序的
 * <p>它是通过{@link #collation}的顺序获取Map中的值
 * <p>所以Excel的列数是根据{@link #collation}确定的（可以过滤不需要的属性）
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @see AbstractWriterResolver
 * @see WriterResolver
 */
public class MapCollationResolver extends AbstractWriterResolver<Map<String, Object>> {
	
	/**
	 * <p>排序规则，对列进行排序，可以指定map中的key写入Excel的顺序
	 * <p>可以过滤不需要的属性
	 */
	private String[] collation;

	/**
	 * <p>Excel写入数据解析器
	 * <p>Map有序数据解析器
	 * <p>此解析器写入Excel的列是有序序的
	 * <p>它是通过{@link #collation}的顺序获取Map中的值
	 * <p>所以Excel的列数是根据{@link #collation}确定的（可以过滤不需要的属性）
	 * @param data 数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，可以过滤不需要的属性
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public MapCollationResolver(List<Map<String, Object>> data, String[] collation) {
		super(data);
		this.collation = collation;
	}

	@Override
	public void title(Row row) {
		for(int i = 0; i < collation.length; i++) {
			row.createCell(i).setCellValue(collation[i]);
		}
	}

	@Override
	public void row(Row row, int index) {
		Map<String, Object> map = get(index);
		for(int i = 0; i < collation.length; i++) {
			cell(row.createCell(i), map.get(collation[i]), index, i);
		}
	}

}
