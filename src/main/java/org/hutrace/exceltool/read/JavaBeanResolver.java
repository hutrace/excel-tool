package org.hutrace.exceltool.read;

import java.beans.IntrospectionException;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;

import org.apache.poi.ss.usermodel.Row;
import org.hutrace.exceltool.annotation.ExcelField;
import org.hutrace.exceltool.exception.ExcelReaderException;
import org.hutrace.exceltool.utils.TypeUtils;

/**
 * <p>Excel读取数据解析器
 * <p>JavaBean解析器
 * <p>将Excel的数据解析成JavaBean
 * <pre>
 *  注意事项：
 *  	获取Excel的第一行作为标题
 *  	可在JavaBean中字段上添加{@link ExcelField}注解，使用title值匹配Excel的标题
 *  	JavaBean只会解析一次，后面每次都按照标题行的长度、顺序执行（也就是说要保证内容行的列数不能超过标题行的列数，超出的列是读不到的。）
 *  </pre>
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @param <T>
 */
public class JavaBeanResolver<T> extends AbstractReaderResolver<Object> {
	
	/**
	 * List中的JavaBean对象
	 */
	private T t;
	
	/**
	 * JavaBean泛型类
	 */
	private Class<T> clazs;
	
	/**
	 * 根据标题按顺序储存的JavaBean中的setter方法数组
	 */
	private Method[] method;
	
	/**
	 * 根据标题按顺序储存的JavaBean的字段类型
	 */
	private Class<?>[] types;
	
	public JavaBeanResolver(Class<T> clazs) {
		this.clazs = clazs;
	}
	
	@Override
	public void title(Row row) {
		Field[] fields = clazs.getDeclaredFields();
		int num = row.getLastCellNum();
		method = new Method[num];
		types = new Class<?>[num];
		Field field;
		ExcelField annot;
		String name;
		for(int i = 0; i < num; i++) {
			for(int j = 0; j < fields.length; j++) {
				field = fields[j];
				annot = field.getAnnotation(ExcelField.class);
				if(annot != null) {
					name = annot.title();
				}else {
					name = field.getName();
				}
				if(name.equals(cell(row.getCell(i), 0, i))) {
					method[i] = method(field.getName());
					types[i] = field.getType();
					break;
				}
			}
		}
	}
	
	/**
	 * <p>构造字段的setter方法
	 * @param name 字段名称
	 * @return 方法({@link Method})
	 */
	private Method method(String name) {
		try {
			return new PropertyDescriptor(name, clazs).getWriteMethod();
		}catch (IntrospectionException e) {
			e.printStackTrace();
			throw new ExcelReaderException(e);
		}
	}
	
	/**
	 * <p>构造JavaBean对象
	 * @return JavaBean对象
	 */
	private T instance() {
		try {
			return clazs.newInstance();
		}catch (Exception e) {
			throw new ExcelReaderException("Failed to create a JavaBean object", e);
		}
	}

	@Override
	public void row(Row row, int index) {
		t = instance();
		for(int i = 0; i < method.length; i++) {
			if(method[i] != null) {
				try {
					method[i].invoke(t, TypeUtils.cast(cell(row.getCell(i), index, i), types[i]));
				}catch (Exception e) {
					throw new ExcelReaderException("Error reading data, row " + index + ", column " + i, e);
				}
			}
		}
		add(t);
	}

}
