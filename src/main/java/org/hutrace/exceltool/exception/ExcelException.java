package org.hutrace.exceltool.exception;

/**
 * <p>Excel异常
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月14日
 */
public class ExcelException extends Exception {

	private static final long serialVersionUID = 1L;
	
	public ExcelException(String message) {
		super(message);
	}

	public ExcelException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelException(Throwable cause) {
		super(cause);
	}
	
}
