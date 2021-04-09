package com.framework.poi.export.excel;

/**
 * <p>
 * 描述: 导出行接口，如果需要自定义每一行的行高，可以实现该接口
 * </p>
 * 
 * @author aLiang
 * @date 2020年8月20日 下午4:15:38
 */
public interface PoiExportRow {

	/**
	 * 
	 * <p>
	 * 行高
	 * </P>
	 * 
	 * @return
	 */
	Integer getHeight();
}
