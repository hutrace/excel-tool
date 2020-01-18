package org.hutrace.exceltool.write;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.hutrace.exceltool.annotation.ExcelField;
import org.hutrace.exceltool.exception.ExcelReaderException;
import org.hutrace.exceltool.exception.ExcelWriteException;

/**
 * <p>Excel写入数据解析器
 * <p>JavaBean数据解析器
 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
 * <p>该解析器是按照{@link #collation}的顺序进行解析，生成Excel列的顺序会按照{@link #collation}的顺序而定
 * <p><b>需要注意：可以通过设置{@link #collation}来对Excel写入列的个数做处理（过滤不需要的字段），该解析器是根据{@link #collation}来获取JavaBean中的数据的</b>
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月15日
 * @see AbstractWriterResolver
 * @see WriterResolver
 */
public class JavaBeanCollationResolver extends AbstractWriterResolver<Object> {
	
	/**
	 * <p>排序规则，对列进行排序，可以指定JavaBean中的字段写入Excel的顺序
	 * <p>也可以过滤不需要的字段
	 */
	private String[] collation;
	
	/**
	 * JavaBean字段的get方法数组
	 */
	private Method[] methods;

	/**
	 * <p>Excel写入数据解析器
	 * <p>JavaBean数据解析器
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该解析器是按照{@link #collation}的顺序进行解析，生成Excel列的顺序会按照{@link #collation}的顺序而定
	 * <p><b>需要注意：可以通过设置{@link #collation}来对Excel写入列的个数做处理（过滤不需要的字段），该解析器是根据{@link #collation}来获取JavaBean中的数据的</b>
	 * @param data 写入Excel的数据
	 * @param collation 排序规则，对列进行排序，可以指定JavaBean中的字段写入Excel的顺序，也可以过滤不需要的字段
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public JavaBeanCollationResolver(List<Object> data, String[] collation) {
		super(data);
		this.collation = collation;
	}

	@Override
	public void title(Row row) {
		Object obj = get(0);
		Class<?> clazs = obj.getClass();
		Field[] fields = clazs.getDeclaredFields();
		if(collation.length > fields.length) {
			throw new ExcelWriteException("The 'collation' length must not be greater than the number of 'fields'");
		}
		methods = new Method[collation.length];
		Field field;
		ExcelField annot;
		String name;
		for(int i = 0; i < collation.length; i++) {
			name = collation[i];
			for(int j = 0; j < fields.length; j++) {
				field = fields[j];
				if(name.equals(field.getName())) {
					methods[i] = method(name, clazs);
					annot = field.getAnnotation(ExcelField.class);
					if(annot != null) {
						name = annot.title();
					}
					break;
				}
			}
			row.createCell(i).setCellValue(name);
		}
	}
	
	/**
	 * <p>构造字段的getter方法
	 * @param name 字段名称
	 * @return 方法({@link Method})
	 */
	private Method method(String name, Class<?> clazs) {
		try {
			return new PropertyDescriptor(name, clazs).getReadMethod();
		}catch (IntrospectionException e) {
			e.printStackTrace();
			throw new ExcelReaderException(e);
		}
	}

	@Override
	public void row(Row row, int index) {
		Object obj = get(index);
		for(int i = 0; i < methods.length; i++) {
			try {
				cell(row.createCell(i), methods[i].invoke(obj), index, i);
			}catch (Exception e) {
				throw new ExcelWriteException("Error writing data, row " + index + ", column " + i, e);
			}
		}
	}

}
