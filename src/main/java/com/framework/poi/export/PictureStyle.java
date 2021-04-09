package com.framework.poi.export;

import java.math.BigDecimal;

import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * <p>
 * 图片样式
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 下午2:25:26
 */
@Data
@EqualsAndHashCode(callSuper = false)
public class PictureStyle extends Style {

	/**
	 * 图片宽度百分比
	 */
	private BigDecimal picWidth;

	/**
	 * 图片高度百分比
	 */
	private BigDecimal picHeight;

	/**
	 * 图片压缩比例
	 */
	private BigDecimal pictureScale;
}
