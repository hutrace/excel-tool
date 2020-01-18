package org.hutrace.exceltool.exception;

/**
 * <p>读取Excel出错的异常
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月14日
 */
public class ExcelReaderException extends RuntimeException {

	private static final long serialVersionUID = 1L;

	public ExcelReaderException(String message) {
		super(message);
	}

	public ExcelReaderException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelReaderException(Throwable cause) {
		super(cause);
	}

}