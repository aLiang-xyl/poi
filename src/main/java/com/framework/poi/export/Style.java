package com.framework.poi.export;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * <p>
 * 样式
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 下午1:50:55
 */
@Data
@Accessors(chain = true)
public class Style {

	/**
	 * 默认字体大小
	 */
	public static final Integer DEFAULT_FONT_SIZE = 13;
	/**
	 * 默认字体
	 */
	public static final String DEFAULT_FONT_TYPE = "仿宋";

	/**
	 * 高度
	 */
	private Integer height;

	/**
	 * 宽度
	 */
	private Integer width;

	/**
	 * 是否加粗
	 */
	private Boolean boldAble;

	/**
	 * 字体类型
	 */
	private String fontType;

	/**
	 * 字体大小
	 */
	private Integer fontSize;

	/**
	 * 字体颜色
	 */
	private String fontColor;

	/**
	 * 
	 * <p>
	 * 居中表格样式
	 * </P>
	 * 
	 * @param width
	 * @return
	 */
//	public static Style tableCenterStyle(Integer width) {
//		return new Style().setAlignment(AlignmentEnum.CENTER).setParagraphAlignment(AlignmentEnum.CENTER)
//				.setVerticalAlignment(AlignmentEnum.CENTER).setWidth(width);
//	}
}
