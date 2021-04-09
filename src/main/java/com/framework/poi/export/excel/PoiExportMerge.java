package com.framework.poi.export.excel;

import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

import lombok.Getter;

/**
 * <p>
 * 描述: 合并信息
 * </p>
 * 
 * @author aLiang
 * @date 2020-12-21 15:46:12
 */
@Getter
public class PoiExportMerge {

	/**
	 * 列序号
	 * 
	 * 
	 */
	private Integer columnIndex;

	/**
	 * 合并信息，key是合并值信息
	 */
	private Map<String, PoiExportMergeRow> mergeMap;

	/**
	 * 
	 * <p>
	 * 添加一个合并项
	 * </P>
	 * 
	 * @param key
	 * @param startRowNumber
	 */
	public void put(Object key, Integer startRowNumber) {
		if (key == null) {
			return;
		}
		PoiExportMergeRow merge = new PoiExportMergeRow();
		
		merge.setStartRowNumber(startRowNumber);
		merge.setEndRowNumber(startRowNumber);
		mergeMap.put(key.toString(), merge);
	}

	/**
	 * 
	 * <p>
	 * 合并项行数加1
	 * </P>
	 * 
	 * @param key
	 */
	public void add(Object key) {
		if (key == null) {
			return;
		}
		PoiExportMergeRow merge = mergeMap.get(key.toString());
		if (merge == null) {
			return;
		}
		Integer endRowNumber = merge.getEndRowNumber();
		endRowNumber++;
		merge.setEndRowNumber(endRowNumber);
	}

	public PoiExportMerge(Integer columnIndex) {
		super();
		this.columnIndex = columnIndex;
		mergeMap = new HashMap<>();
	}
	
	public Collection<PoiExportMergeRow> getAllMerge() {
		return mergeMap.values();
	}
}
