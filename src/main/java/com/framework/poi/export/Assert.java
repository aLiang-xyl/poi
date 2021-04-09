package com.framework.poi.export;

import java.util.Collection;
import java.util.Map;

/**
 * <p>
 * 描述: 断言工具类，用来校验参数
 * </p>
 * 
 * @author aLiang
 * @date 2020年8月26日 上午11:40:51
 */
public class Assert {

	/**
	 * 如果expression是true，则校验通过
	 * 
	 * @param expression
	 * @param message
	 * @throws BusinessException 如果expression为false
	 */
	public static void isTrue(boolean expression, String message) {
		if (!expression) {
			throw new BusinessException(message);
		}
	}

	/**
	 * 如果expression是false，则校验通过
	 * 
	 * @param expression
	 * @param message
	 * @throws BusinessException 如果expression为true
	 */
	public static void isFalse(boolean expression, String message) {
		if (expression) {
			throw new BusinessException(message);
		}
	}

	/**
	 * 如果object不是null则校验通过
	 * 
	 * @param object
	 * @param message
	 * @throws BusinessException 如果object为null
	 */
	public static void notNull(Object object, String message) {
		if (object == null) {
			throw new BusinessException(message);
		}
	}

	/**
	 * 如果collection不为null并且size大于0则校验通过
	 * 
	 * @param collection
	 * @param message
	 * @throws BusinessException 如果collection为null并且size为0
	 */
	public static void notEmpty(Collection<?> collection, String message) {
		if (collection == null || collection.isEmpty()) {
			throw new BusinessException(message);
		}
	}

	/**
	 * 如果map不为null并且size大于0则校验通过
	 * 
	 * @param map
	 * @param message
	 * @throws BusinessException 如果map为null并且size为0
	 */
	public static void notEmpty(Map<?, ?> map, String message) {
		if (map == null || map.isEmpty()) {
			throw new BusinessException(message);
		}
	}

	/**
	 * 如果str不为null并且长度大于0则校验通过
	 * 
	 * @param str
	 * @param message
	 * @throws BusinessException 如果str为null并且长度为0
	 */
	public static void notEmpty(String str, String message) {
		if (str == null || "".equals(str)) {
			throw new BusinessException(message);
		}
	}
}
