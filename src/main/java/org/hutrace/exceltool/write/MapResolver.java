package org.hutrace.exceltool.write;

import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;

/**
 * <p>Excel写入数据解析器
 * <p>Map数据解析器
 * <p>此解析器写入Excel的列是无序的（针对整体而言无序）
 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
 * <p>你不能确定哪一列排在第一，哪一列排在最后。
 * <p>如果你想要列有序，使用{@link MapCollationResolver}
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月15日
 * @see AbstractWriterResolver
 * @see WriterResolver
 */
public class MapResolver extends AbstractWriterResolver<Map<String, Object>> {
	
	/**
	 * <p>Excel写入数据解析器
	 * <p>Map数据解析器
	 * <p>此解析器写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * <p>如果你想要列有序，使用{@link MapCollationResolver}
	 * @param data 写入Excel的数据
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public MapResolver(List<Map<String, Object>> data) {
		super(data);
	}

	@Override
	public void title(Row row) {
		Map<String, Object> map = get(0);
		Set<String> keys = map.keySet();
		int i = 0;
		for(String key : keys) {
			row.createCell(i++).setCellValue(key);
		}
	}

	@Override
	public void row(Row row, int index) {
		Collection<Object> values = get(index).values();
		int i = 0;
		for(Object value : values) {
			cell(row.createCell(i), value, index, i++);
		}
	}

}
