package com.framework.poi.export;

import java.beans.PropertyDescriptor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.TemporalAccessor;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.BeanUtils;

import com.alibaba.fastjson.JSONObject;

import lombok.extern.log4j.Log4j2;

/**
 * <p>
 * 工具类
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 下午6:35:41
 */
@Log4j2
public class PoiUtils {

	/**
	 * 日期时间正则表达式
	 */
	private static final String LOCAL_DATE_TIME_PATTERN = "[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}.{0,1}[0-9]{0,4}";
	/**
	 * 日期正则表达式
	 */
	private static final String LOCAL_DATE_PATTERN = "[0-9]{4}-[0-9]{2}-[0-9]{2}";
	/**
	 * 时间正则表达式
	 */
	private static final String LOCAL_TIME_PATTERN = "[0-9]{2}:[0-9]{2}:[0-9]{2}.{0,1}[0-9]{0,4}";

	/**
	 * 
	 * <p>
	 * 从对象中取值
	 * </P>
	 * 
	 * @param rowObj
	 * @param fieldName
	 * @param isJson
	 * @return
	 */
	public static Object getValue(Object rowObj, String fieldName, boolean isJson) {
		if (rowObj == null) {
			return null;
		}
		Object value = rowObj;
		String[] fieldArr = fieldName.split("\\.");
		int length = fieldArr.length;
		for (int i = 0; i < length; i++) {
			String field = fieldArr[i];
			if (isJson) {
				if (i < length - 1) {
					JSONObject json = (JSONObject) value;
					value = json.getJSONObject(field);
					if (value == null) {
						return null;
					}
				} else {
					JSONObject json = (JSONObject) value;
					return json.get(field);
				}
			} else {
				PropertyDescriptor pd = BeanUtils.getPropertyDescriptor(value.getClass(), field);
				if (pd == null) {
					log.error("{}未查到成员变量或者方法{}", rowObj.getClass(), fieldName);
					return null;
				}
				Method method = pd.getReadMethod();
				try {
					value = method.invoke(value);
					if (value == null) {
						return null;
					}
				} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
					log.error("读取数据出错", e);
				}
			}
		}
		return value;
	}

	/**
	 * 
	 * <p>
	 * 值格式化
	 * </P>
	 * 
	 * @param value
	 * @param dateTimeFromat
	 * @param valueExchange
	 * @param isJson
	 * @return
	 */
	public static Object formatValue(Object value, String dateTimeFromat, String valueExchange, boolean isJson) {
		if (value == null) {
			return "";
		}

		boolean dateTimeFromatIsNotNull = !StringUtils.isEmpty(dateTimeFromat);

		if (isJson && dateTimeFromatIsNotNull) {
			// 如果数据源是json格式的，则需要转成日期格式
			String date = value.toString();
			if (Pattern.matches(LOCAL_DATE_TIME_PATTERN, date)) {
				value = LocalDateTime.parse(date);
			} else if (Pattern.matches(LOCAL_DATE_PATTERN, date)) {
				value = LocalDate.parse(date);
			} else if (Pattern.matches(LOCAL_TIME_PATTERN, date)) {
				value = LocalTime.parse(date);
			}
		}

		if (dateTimeFromatIsNotNull
				&& (value instanceof LocalDate || value instanceof LocalDateTime || value instanceof LocalTime)) {
			// 原始数据是日期格式的或者通过json转换成日期格式的，则进行格式化
			DateTimeFormatter pattern = DateTimeFormatter.ofPattern(dateTimeFromat);
			return pattern.format((TemporalAccessor) value);
		}

		if (!StringUtils.isEmpty(valueExchange)) {
			JSONObject json = JSONObject.parseObject(valueExchange);
			Object object = json.get(value.toString());
			if (object != null) {
				return object;
			}
		}

		return value;
	}
}
