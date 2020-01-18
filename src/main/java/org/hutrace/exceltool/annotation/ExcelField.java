package org.hutrace.exceltool.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * <p>Excel字段注解
 * <p>通过此注解可以指定读写时的标题
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月14日
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelField {
	
	/**
	 * Excel的标题
	 * @return
	 */
	String title();
	
}
