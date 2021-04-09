package com.framework.poi.export;

import lombok.Data;

/**
 * <p>
 * 列合并信息
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 上午11:03:54
 */
@Data
public class ColumnMerge {

	/**
	 * 合并开始列
	 */
	private Integer mergeColumnStart;

	/**
	 * 合并结束列
	 */
	private Integer mergeColumnEnd;

	public ColumnMerge(Integer mergeColumnStart, Integer mergeColumnEnd) {
		super();
		this.mergeColumnStart = mergeColumnStart;
		this.mergeColumnEnd = mergeColumnEnd;
	}

	public ColumnMerge() {
		super();
	}
}
