package com.framework.poi.export.excel.vo;

import com.framework.poi.export.excel.entity.ExcelColumn;
import io.swagger.annotations.ApiModel;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * <p>
 * excel表格列配置
 * </p>
 *
 * @author aLiang
 * @since 2021-01-13
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
@ApiModel(value="ExcelColumnVo对象", description="excel表格列配置")
public class ExcelColumnVo extends ExcelColumn {

    private static final long serialVersionUID=1L;

}
