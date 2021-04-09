package com.framework.poi.export.excel;

import lombok.Data;

/**
 * <p>
 * 描述: 合并行信息
 * </p>
 * 
 * @author aLiang
 * @date 2020-12-22 15:07:44
 */
@Data
public class PoiExportMergeRow {

	/**
	 * 合并开始行
	 * 
	 * 
	 */
	private Integer startRowNumber;

	/**
	 * 合并结束行
	 * 
	 * 
	 */
	private Integer endRowNumber;

}
