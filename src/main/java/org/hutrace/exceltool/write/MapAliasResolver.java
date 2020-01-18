package org.hutrace.exceltool.write;

import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.hutrace.exceltool.pojo.TitleAlias;

/**
 * <p>Excel写入数据解析器
 * <p>Map别名数据解析器
 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
 * <p>此解析器写入Excel的列是无序的（针对整体而言无序）
 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
 * <p>你不能确定哪一列排在第一，哪一列排在最后。
 * <p>如果你想要列有序，使用{@link MapAliasCollationResolver}
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月15日
 * @see AbstractWriterResolver
 * @see WriterResolver
 */
public class MapAliasResolver extends AbstractWriterResolver<Map<String, Object>> {
	
	/**
	 * <p>标题别名数组
	 * <p>通过它可以设置写入Excel的标题内容
	 */
	private TitleAlias[] titleAlias;

	/**
	 * <p>Excel写入数据解析器
	 * <p>Map别名数据解析器
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此解析器写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * <p>如果你想要列有序，使用{@link MapAliasCollationResolver}
	 * @param data 写入Excel的数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public MapAliasResolver(List<Map<String, Object>> data, TitleAlias[] titleAlias) {
		super(data);
		this.titleAlias = titleAlias;
	}

	@Override
	public void title(Row row) {
		Map<String, Object> map = get(0);
		Set<String> keys = map.keySet();
		int i = 0;
		for(String key : keys) {
			for(int j = 0; j < titleAlias.length; j++) {
				if(key.equals(titleAlias[j].getAlias())) {
					key = titleAlias[j].getTitle();
					break;
				}
			}
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
