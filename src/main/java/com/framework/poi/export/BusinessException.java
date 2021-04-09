package com.framework.poi.export;




/**
 * @author xingyl
 * @date 2019年4月22日下午2:11:23
 */
public class BusinessException extends RuntimeException {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;


	public BusinessException() {
	}

	public BusinessException(Throwable ex) {
		super(ex);
	}

	public BusinessException(String message) {
		super(message);
	}


	public BusinessException(String message, Throwable ex) {
		super(message, ex);
	}
}
