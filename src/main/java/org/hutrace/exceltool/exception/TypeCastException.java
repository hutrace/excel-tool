package org.hutrace.exceltool.exception;

/**
 * <p>类型转换异常
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月15日
 */
public class TypeCastException extends RuntimeException {

	private static final long serialVersionUID = 1L;

	public TypeCastException(String message) {
		super(message);
	}

	public TypeCastException(String message, Throwable cause) {
		super(message, cause);
	}

	public TypeCastException(Throwable cause) {
		super(cause);
	}

}