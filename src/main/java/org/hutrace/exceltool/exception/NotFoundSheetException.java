package org.hutrace.exceltool.exception;

/**
 * <p>没有找到Excel的sheet抛出的异常
 * <p>此异常继承{@link RuntimeException}
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @see RuntimeException
 */
public class NotFoundSheetException extends RuntimeException {

	private static final long serialVersionUID = 1L;

	public NotFoundSheetException(String message) {
		super(message);
	}

	public NotFoundSheetException(String message, Throwable cause) {
		super(message, cause);
	}

	public NotFoundSheetException(Throwable cause) {
		super(cause);
	}

}
