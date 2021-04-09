package com.framework.poi.export.excel.vo;

import java.util.List;

import com.framework.poi.export.excel.entity.ExcelColumn;
import com.framework.poi.export.excel.entity.ExcelExportSheet;
import com.framework.poi.export.excel.entity.ExcelExportTemplate;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * <p>
 * excel导出sheet配置
 * </p>
 *
 * @author aLiang
 * @since 2021-01-13
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
@ApiModel(value = "ExcelExportSheetVo对象", description = "excel导出sheet配置")
public class ExcelExportSheetVo extends ExcelExportSheet {

	private static final long serialVersionUID = 1L;

	@ApiModelProperty(value = "列信息")
	private List<ExcelColumn> columns;
	
    @ApiModelProperty(value = "excel导出模板信息")
    private ExcelExportTemplate excelExportTemplate;
}
