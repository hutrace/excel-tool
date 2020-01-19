package org.hutrace.exceltool.write;

import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.hutrace.exceltool.pojo.TitleAlias;

/**
 * <p>Excel写入数据解析器
 * <p>Map别名有序数据解析器
 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
 * <p>此解析器写入Excel的列是有序序的
 * <p>它是通过{@link #collation}的顺序获取Map中的值
 * <p>所以Excel的列数是根据{@link #collation}确定的（可以过滤不需要的属性）
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @see AbstractWriterResolver
 * @see WriterResolver
 */
public class MapAliasCollationResolver extends AbstractWriterResolver<Map<String, Object>> {
	
	/**
	 * <p>标题别名数组
	 * <p>通过它可以设置写入Excel的标题内容
	 */
	private TitleAlias[] titleAlias;
	
	/**
	 * <p>排序规则，对列进行排序，可以指定map中的key写入Excel的顺序
	 * <p>也可以过滤不需要的属性
	 */
	private String[] collation;

	/**
	 * <p>Excel写入数据解析器
	 * <p>Map别名有序数据解析器
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此解析器写入Excel的列是有序序的
	 * <p>它是通过{@link #collation}的顺序获取Map中的值
	 * <p>所以Excel的列数是根据{@link #collation}确定的（可以过滤不需要的属性）
	 * @param data 写入Excel的数据
	 * @param titleAlias Excel的标题别名
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public MapAliasCollationResolver(List<Map<String, Object>> data, TitleAlias[] titleAlias, String[] collation) {
		super(data);
		this.titleAlias = titleAlias;
		this.collation = collation;
	}

	@Override
	public void title(Row row) {
		String title;
		for(int i = 0; i < collation.length; i++) {
			title = collation[i];
			for(int j = 0; j < titleAlias.length; j++) {
				if(title.equals(titleAlias[j].getAlias())) {
					title = titleAlias[j].getTitle();
					break;
				}
			}
			row.createCell(i).setCellValue(title);
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
