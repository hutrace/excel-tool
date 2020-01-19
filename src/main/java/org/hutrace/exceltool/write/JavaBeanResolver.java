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
 * <p>该解析器是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
 * <p>当然，你可以使用{@link JavaBeanCollationResolver}解析器，它可以自定义顺序
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @see AbstractWriterResolver
 * @see WriterResolver
 */
public class JavaBeanResolver extends AbstractWriterResolver<Object> {
	
	/**
	 * JavaBean字段的get方法数组
	 */
	private Method[] methods;

	/**
	 * <p>Excel写入数据解析器
	 * <p>JavaBean数据解析器
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该解析器是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * <p>当然，你可以使用{@link JavaBeanCollationResolver}解析器，它可以自定义顺序
	 * @param data 写入Excel的数据
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public JavaBeanResolver(List<Object> data) {
		super(data);
	}

	@Override
	public void title(Row row) {
		Object obj = get(0);
		Class<?> clazs = obj.getClass();
		Field[] fields = clazs.getDeclaredFields();
		methods = new Method[fields.length];
		Field field;
		ExcelField annot;
		String name;
		for(int i = 0; i < fields.length; i++) {
			field = fields[i];
			name = field.getName();
			methods[i] = method(name, clazs);
			annot = field.getAnnotation(ExcelField.class);
			if(annot != null) {
				name = annot.title();
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
