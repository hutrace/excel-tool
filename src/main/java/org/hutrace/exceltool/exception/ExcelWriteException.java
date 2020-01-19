package org.hutrace.exceltool.exception;

/**
 * <p>写入Excel出错的异常
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 */
public class ExcelWriteException extends RuntimeException {

	private static final long serialVersionUID = 1L;

	public ExcelWriteException(String message) {
		super(message);
	}

	public ExcelWriteException(String message, Throwable cause) {
		super(message, cause);
	}

	public ExcelWriteException(Throwable cause) {
		super(cause);
	}

}