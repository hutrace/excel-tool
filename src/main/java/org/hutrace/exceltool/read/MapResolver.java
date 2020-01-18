package org.hutrace.exceltool.read;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;

/**
 * <p>读取Excel数据解析器
 * <p>将数据解析成Map类型
 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
 * <p>Excel的标题作为Map的键
 * <p>Map的长度(size)也将会使用标题的个数来确定
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月14日
 */
public class MapResolver extends AbstractReaderResolver<Map<String, Object>> {

	/**
	 * Map的key数组
	 */
	private String[] mapKeys;
	
	/**
	 * List中的Map全局缓存
	 */
	private Map<String, Object> map;
	
	@Override
	public void title(Row row) {
		int cellNum = row.getLastCellNum();
		mapKeys = new String[cellNum];
		for(int i = 0; i < cellNum; i++) {
			mapKeys[i] = cell(row.getCell(i), 0, i).toString();
		}
	}

	@Override
	public void row(Row row, int index) {
		map = new HashMap<>(mapKeys.length);
		for(int i = 0; i < mapKeys.length; i++) {
			map.put(mapKeys[i], cell(row.getCell(i), index, i));
		}
		add(map);
	}


}
