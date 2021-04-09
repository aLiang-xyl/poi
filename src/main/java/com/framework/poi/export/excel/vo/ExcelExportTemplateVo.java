package com.framework.poi.export.excel.vo;

import com.framework.poi.export.excel.entity.ExcelExportTemplate;
import io.swagger.annotations.ApiModel;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * <p>
 * excel导出模板
 * </p>
 *
 * @author aLiang
 * @since 2021-01-13
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
@ApiModel(value="ExcelExportTemplateVo对象", description="excel导出模板")
public class ExcelExportTemplateVo extends ExcelExportTemplate {

    private static final long serialVersionUID=1L;

}
