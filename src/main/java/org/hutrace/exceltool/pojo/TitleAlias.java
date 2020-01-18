package org.hutrace.exceltool.pojo;

/**
 * <p>Excel标题与Java字段(属性名)
 * <p>在读取的时候，它会使用{@link #title}匹配Excel的标题，使用{@link #alias}作为Map的key、或者是与JavaBean的字段名匹配
 * <p>在写入的时候，它会使用{@link #alias}匹配Map的key、或者JavaBean的字段名，使用{@link #title}作为Excel的标题
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月14日
 */
public class TitleAlias {
	
	/**
	 * Excel的标题
	 */
	private String title;
	
	/**
	 * Map或JavaBean的属性(字段)名称
	 */
	private String alias;
	
	private TitleAlias() {}
	
	/**
	 * 静态构造{@link TitleAlias}对象
	 * @param title Excel的标题
	 * @param alias Map或JavaBean的属性(字段)名称
	 * @return {@link TitleAlias}
	 */
	public static TitleAlias build(String title, String alias) {
		TitleAlias titleAlias = new TitleAlias();
		titleAlias.title = title;
		titleAlias.alias = alias;
		return titleAlias;
	}
	
	/**
	 * 获取Excel的标题
	 * @return Excel的标题
	 */
	public String getTitle() {
		return title;
	}
	
	/**
	 * 获取Map或JavaBean的属性(字段)名称
	 * @return Map或JavaBean的属性(字段)名称
	 */
	public String getAlias() {
		return alias;
	}
	
}
