package com.framework.poi.export.word;

import java.util.LinkedList;
import java.util.List;

import com.framework.poi.export.Assert;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * <p>
 * word内容接口，其实现可以是一行文本内容，也可以是一个表格内容</br>
 * 所有的内容可以看成一块一块的，每一块中分为行，每一行中分为列，每一列中又分为段落
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月19日 下午2:06:12
 */
@Data
@Accessors(chain = true)
public class DocContent {

	/**
	 * 
	 * <p>
	 * 列索引
	 * </P>
	 * 
	 */
	private Integer contentIndex;

	public DocContent() {
		super();
	}

	public DocContent(Integer contentIndex) {
		super();
		this.contentIndex = contentIndex;
	}

	/**
	 * 
	 * <p>
	 * 添加元素
	 * </P>
	 * 
	 * @param <T>
	 * @param t
	 * @return
	 */
	@SuppressWarnings("unchecked")
	public <T> List<T> addContent(T... t) {
		Assert.isTrue(t != null && t.length > 0, "参数错误");
		List<T> list = new LinkedList<T>();
		for (T t2 : t) {
			list.add(t2);
		}
		return list;
	}
}
