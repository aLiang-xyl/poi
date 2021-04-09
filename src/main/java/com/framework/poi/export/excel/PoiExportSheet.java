package com.framework.poi.export.excel;

import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.util.CollectionUtils;

import com.alibaba.fastjson.JSONArray;
import com.framework.poi.export.Assert;
import com.framework.poi.export.excel.entity.ExcelColumn;
import com.framework.poi.export.excel.vo.ExcelExportSheetVo;

import lombok.Data;
import lombok.EqualsAndHashCode;

/**
 * <p>
 * 描述:
 * </p>
 * 
 * @author aLiang
 * @date 2020-12-21 15:46:12
 */
@Data
@EqualsAndHashCode(callSuper=false)
public class PoiExportSheet extends ExcelExportSheetVo {

	/**   
	 * @Fields serialVersionUID 
	 */
	private static final long serialVersionUID = 1L;

	/**
	 * 数据是否是json格式
	 * @see JSONArray
	 */
	private boolean rowsIsJson;
	/**
	 * 所有行数据
	 */
	private Collection<?> rows;

	/**
	 * 合并项信息
	 */
	private Map<Integer, PoiExportMerge> mergeMap;
	
	public PoiExportSheet() {
		super();
		mergeMap = new HashMap<>();
	}
	
	/**
	 * 
	 * <p>
	 * 根据列编号获取合并信息
	 * </P>
	 * @param columnIndex
	 * @return
	 */
	public PoiExportMerge getMerge(Integer columnIndex) {
		PoiExportMerge exportMerge = mergeMap.get(columnIndex);
		if (exportMerge == null) {
			exportMerge = new PoiExportMerge(columnIndex);
			mergeMap.put(columnIndex, exportMerge);
		}
		return exportMerge;
	}
	
	public void putMergeEndRow(Integer columnIndex, Object key, Integer startRowNumber) {
		PoiExportMerge exportMerge = getMerge(columnIndex);
		exportMerge.put(key, startRowNumber);
	} 
	
	public void addMergeEndRow(Integer columnIndex, Object key) {
		PoiExportMerge exportMerge = mergeMap.get(columnIndex);
		exportMerge.add(key);
	}
	
	public boolean containsMerge(Integer columnIndex, Object key) {
		if (key == null) {
			return false;
		}
		PoiExportMerge exportMerge = mergeMap.get(columnIndex);
		if (exportMerge == null) {
			return false;
		}
		Map<String, PoiExportMergeRow> map = exportMerge.getMergeMap();
		return map.containsKey(key.toString());
	}
	
	public void validate() {
		List<ExcelColumn> columns = getColumns();
		Assert.isTrue(!CollectionUtils.isEmpty(columns), "模板数据错误");
		Integer height = getHeight();
		Assert.isTrue(height != null && height > 0, "行高参数错误");
	}
	
}
