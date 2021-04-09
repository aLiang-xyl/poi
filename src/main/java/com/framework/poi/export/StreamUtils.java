package com.framework.poi.export;

import java.math.BigDecimal;
import java.util.Collection;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.function.ToIntFunction;
import java.util.function.ToLongFunction;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * java.util.stream工具类
 * 
 * @author xing
 *
 */
public class StreamUtils {

	/**
	 * 将一个对象列表中的某个属性转换成list并根据条件过滤
	 * 
	 * @param <T>
	 * @param <R>
	 * @param list
	 * @param mapper
	 * @param predicate
	 * @return
	 */
	public static <T, R> List<R> toListBeforeFilter(Collection<T> list, Function<? super T, ? extends R> mapper,
			Predicate<? super T> predicate) {
		return stream(list).filter(predicate).map(mapper).collect(Collectors.toList());
	}

	/**
	 * 将一个对象列表中的某个属性转换成list并根据条件过滤
	 * 
	 * @param <T>
	 * @param <R>
	 * @param list
	 * @param mapper
	 * @param predicate
	 * @return
	 */
	public static <T, R> List<R> toListBeforeFilterAndDistinct(Collection<T> list,
			Function<? super T, ? extends R> mapper, Predicate<? super T> predicate) {
		return stream(list).filter(predicate).map(mapper).distinct().collect(Collectors.toList());
	}

	/**
	 * 将一个对象列表中的某个属性转换成list并根据条件过滤
	 * 
	 * @param <T>
	 * @param <R>
	 * @param list
	 * @param mapper
	 * @param predicate
	 * @return
	 */
	public static <T, R> List<R> toListAfterFilter(Collection<T> list, Function<? super T, ? extends R> mapper,
			Predicate<? super R> predicate) {
		return stream(list).map(mapper).filter(predicate).collect(Collectors.toList());
	}

	/**
	 * 将一个对象列表中的某个属性转换成list并根据条件过滤
	 * 
	 * @param <T>
	 * @param <R>
	 * @param list
	 * @param mapper
	 * @param predicate
	 * @return
	 */
	public static <T, R> List<R> toListAfterFilterAndDistinct(Collection<T> list,
			Function<? super T, ? extends R> mapper, Predicate<? super R> predicate) {
		return stream(list).map(mapper).filter(predicate).distinct().collect(Collectors.toList());
	}

	/**
	 * 将一个对象列表中的某个属性转换成list
	 * 
	 * @param <T>
	 * @param <R>
	 * @param list
	 * @param mapper
	 * @return
	 */
	public static <T, R> List<R> toList(Collection<T> list, Function<? super T, ? extends R> mapper) {
		return stream(list).filter(v -> v != null).map(mapper).filter(v -> v != null).collect(Collectors.toList());
	}

	/**
	 * 将一个对象列表中的某个属性转换成list并去重
	 * 
	 * @param <T>
	 * @param <R>
	 * @param list
	 * @param mapper
	 * @return
	 */
	public static <T, R> List<R> toListAndDistinct(Collection<T> list, Function<? super T, ? extends R> mapper) {
		return stream(list).filter(v -> v != null).map(mapper).filter(v -> v != null).distinct()
				.collect(Collectors.toList());
	}

	/**
	 * 将一个列表按某个属性分组
	 * 
	 * @param <T>
	 * @param <R>
	 * @param list
	 * @param mapper
	 * @return
	 */
	public static <T, R> Map<R, List<T>> groupingBy(Collection<T> list, Function<? super T, ? extends R> mapper) {
		return stream(list).collect(Collectors.groupingBy(mapper));
	}

	/**
	 * 分组后取list中的第一条,一般用来分组对应的key只有一条数据
	 * 
	 * @param <T>
	 * @param <R>
	 * @param list
	 * @param mapper
	 * @return
	 */
	public static <T, R> Map<R, T> groupingByGetOne(Collection<T> list, Function<? super T, ? extends R> mapper) {
		Map<? extends R, List<T>> map = stream(list).collect(Collectors.groupingBy(mapper));
		HashMap<R, T> hashMap = new HashMap<>();
		map.entrySet().forEach(entry -> {
			if (entry.getValue() != null && entry.getValue().size() > 0) {
				hashMap.put(entry.getKey(), entry.getValue().get(0));
			}
		});
		// 也可以这样写
		// list.stream().collect(Collectors.toMap(mapper, Function.identity(), (key,
		// value) -> value))
		return hashMap;
	}


	/**
	 * 计算总数
	 * 
	 * @param <T>
	 * @param list
	 * @param mapper
	 * @return
	 */
	public static <T> int sumInt(Collection<T> list, ToIntFunction<? super T> mapper) {
		return stream(list).mapToInt(mapper).sum();
	}

	/**
	 * 计算总数
	 * 
	 * @param <T>
	 * @param list
	 * @param mapper
	 * @return
	 */
	public static <T> long sumLong(Collection<T> list, ToLongFunction<? super T> mapper) {
		return stream(list).mapToLong(mapper).sum();
	}

	/**
	 * 过滤后计算总数
	 * 
	 * @param <T>
	 * @param list
	 * @param mapper
	 * @param predicate
	 * @return
	 */
	public static <T> int sumIntFilter(Collection<T> list, ToIntFunction<? super T> mapper,
			Predicate<? super T> predicate) {
		return stream(list).filter(predicate).mapToInt(mapper).sum();
	}

	/**
	 * 过滤后计算总数
	 * 
	 * @param <T>
	 * @param list
	 * @param mapper
	 * @param predicate
	 * @return
	 */
	public static <T> long sumLongFilter(Collection<T> list, ToLongFunction<? super T> mapper,
			Predicate<? super T> predicate) {
		return stream(list).filter(predicate).mapToLong(mapper).sum();
	}

	/**
	 * 
	 * <p>
	 * 计算BigDecimal总数
	 * </P>
	 * 
	 * @param <T>
	 * @param list
	 * @param mapper
	 * @param defaultValue 如果没有值则使用默认值
	 * @return
	 */
	public static <T> BigDecimal sumBigDecimal(List<T> list, Function<? super T, BigDecimal> mapper,
			String defaultValue) {
		Optional<BigDecimal> optional = sumBigDecimal(list, mapper);
		if (!optional.isPresent()) {
			return new BigDecimal(defaultValue);
		}
		return optional.get();
	}

	/**
	 * 
	 * <p>
	 * 计算BigDecimal总数
	 * </P>
	 * 
	 * @param <T>
	 * @param list
	 * @param mapper
	 * @return
	 */
	public static <T> Optional<BigDecimal> sumBigDecimal(List<T> list, Function<? super T, BigDecimal> mapper) {
		Optional<BigDecimal> optional = stream(list).filter(v -> v != null).map(mapper).filter(v -> v != null)
				.reduce((sum, item) -> sum.add(item));
		return optional;
	}

	/**
	 * 过滤数据
	 * 
	 * @param <T>
	 * @param list
	 * @param predicate
	 * @return
	 */
	public static <T> List<T> filter(Collection<T> list, Predicate<? super T> predicate) {
		return stream(list).filter(predicate).collect(Collectors.toList());
	}

	/**
	 * 根据条件查询任意一个查询
	 * 
	 * @param <T>
	 * @param list
	 * @param predicate
	 * @return
	 */
	public static <T> Optional<T> findAny(Collection<T> list, Predicate<? super T> predicate) {
		return stream(list).filter(predicate).findAny();
	}

	/**
	 * 顺序
	 * 
	 * @param <T>
	 * @param <U>
	 * @param list
	 * @param keyExtractor
	 * @return
	 */
	public static <T, U extends Comparable<? super U>> List<T> sortAsc(Collection<T> list,
			Function<? super T, ? extends U> keyExtractor) {
		return sort(list, keyExtractor, false);
	}

	/**
	 * 倒序
	 * 
	 * @param <T>
	 * @param <U>
	 * @param list
	 * @param keyExtractor
	 * @return
	 */
	public static <T, U extends Comparable<? super U>> List<T> sortDesc(Collection<T> list,
			Function<? super T, ? extends U> keyExtractor) {
		return sort(list, keyExtractor, true);
	}

	/**
	 * 去重
	 * 
	 * @param <T>
	 * @param list
	 * @return
	 */
	public static <T> List<T> distinct(Collection<T> list) {
		return stream(list).distinct().collect(Collectors.toList());
	}

	/**
	 * 字符串集合插入字符
	 * 
	 * @param list
	 * @param delimiter
	 * @return
	 */
	public static String joining(Collection<String> list, String delimiter) {
		return stream(list).collect(Collectors.joining(delimiter));
	}

	/**
	 * 插入字符串
	 * 
	 * @param <T>
	 * @param <R>
	 * @param list
	 * @param mapper
	 * @param delimiter
	 * @return String
	 */
	public static <T, R> String joining(Collection<T> list, Function<? super T, ? extends String> mapper,
			String delimiter) {
		return stream(list).filter(v -> v != null).map(mapper).collect(Collectors.joining(delimiter));
	}

	/**
	 * 字符串集合拼接
	 * 
	 * @param list
	 * @return
	 */
	public static String joining(Collection<String> list) {
		return stream(list).collect(Collectors.joining());
	}

	/**
	 * 插入字符串
	 * 
	 * @param <T>
	 * @param <R>
	 * @param list
	 * @param mapper
	 * @return String
	 */
	public static <T, R> String joining(Collection<T> list, Function<? super T, ? extends R> mapper) {
		return stream(list).filter(v -> v != null).map(String::valueOf).collect(Collectors.joining());
	}

	/**
	 * 转换成String List
	 * 
	 * @param list
	 * @return List<String>
	 */
	public static List<String> toStringList(Collection<?> list) {
		return stream(list).filter(v -> v != null).map(String::valueOf).collect(Collectors.toList());
	}

	/**
	 * 排序
	 * 
	 * @param <T>
	 * @param <U>
	 * @param list
	 * @param keyExtractor
	 * @param reversed     是否倒序
	 * @return
	 */
	private static <T, U extends Comparable<? super U>> List<T> sort(Collection<T> list,
			Function<? super T, ? extends U> keyExtractor, boolean reversed) {
		Comparator<? super T> comparator = Comparator.comparing(keyExtractor);
		if (reversed) {
			comparator.reversed();
		}
		return stream(list).sorted(comparator).collect(Collectors.toList());
	}

	/**
	 * list先校验再转换成stream
	 * 
	 * @param <T>
	 * @param <R>
	 * @param list
	 * @return
	 */
	private static <T> Stream<T> stream(Collection<T> list) {
		if (list == null || list.size() == 0) {
			return Stream.empty();
		}
		return list.stream();
	}
}
