package com.framework.poi.export;

import lombok.Data;

/**
 * <p>
 * 合并信息
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 上午11:03:54
 */
@Data
public class RowMerge {

	/**
	 * 列索引
	 */
	private Integer columnIndex;

	/**
	 * 合并开始行
	 */
	private Integer mergeRowStart;

	/**
	 * 合并结束行
	 */
	private Integer mergeRowEnd;

	public RowMerge() {
		super();
	}

	public RowMerge(Integer columnIndex, Integer mergeRowStart, Integer mergeRowEnd) {
		super();
		this.columnIndex = columnIndex;
		this.mergeRowStart = mergeRowStart;
		this.mergeRowEnd = mergeRowEnd;
	}
}
