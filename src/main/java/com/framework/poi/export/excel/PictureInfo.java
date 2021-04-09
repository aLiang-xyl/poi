package com.framework.poi.export.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * 
 * <p>
 * 图片信息
 * </P>
 * 
 * @author aLiang
 * @date 2020年11月23日 下午3:50:18
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class PictureInfo {
 
	/**
	 * 图片路径
	 */
	private String path;
	
	/**
	 * 图片描述
	 */
	private String description;
	
}
