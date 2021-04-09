package com.framework.poi.export;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * <p>
 * word表格样式
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 下午1:50:55
 */
@Data
@Accessors(chain = true)
public class ColumnStyle {

	/**
	 * 表格宽度
	 */
	private Integer width;

	/**
	 * 左右对齐方式
	 */
	private AlignmentEnum alignment;

	/**
	 * 垂直对齐方式
	 */
	private AlignmentEnum verticalAlignment;

	/**
	 * 背景颜色
	 */
	private String backgroundColor;

	public ColumnStyle(Integer width, AlignmentEnum alignment, AlignmentEnum verticalAlignment,
			String backgroundColor) {
		super();
		this.width = width;
		this.alignment = alignment;
		this.verticalAlignment = verticalAlignment;
		this.backgroundColor = backgroundColor;
	}

	public ColumnStyle() {
		super();
	}

	/**
	 * 
	 * <p>
	 * 居中显示
	 * </P>
	 * @param width
	 * @param backgroundColor
	 * @return
	 */
	public static ColumnStyle centerStyle(Integer width, String backgroundColor) {
		return new ColumnStyle(width, AlignmentEnum.CENTER, AlignmentEnum.CENTER, backgroundColor);
	}
	
	/**
	 * 
	 * <p>
	 * 居中显示
	 * </P>
	 * @param width
	 * @return
	 */
	public static ColumnStyle centerStyle(Integer width) {
		return new ColumnStyle(width, AlignmentEnum.CENTER, AlignmentEnum.CENTER, null);
	}
}
