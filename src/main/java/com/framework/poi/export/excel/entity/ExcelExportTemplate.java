package com.framework.poi.export.excel.entity;

import java.io.Serializable;
import java.time.LocalDateTime;

import org.springframework.format.annotation.DateTimeFormat;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.databind.annotation.JsonDeserialize;
import com.fasterxml.jackson.databind.annotation.JsonSerialize;
import com.fasterxml.jackson.datatype.jsr310.deser.LocalDateTimeDeserializer;
import com.fasterxml.jackson.datatype.jsr310.ser.LocalDateTimeSerializer;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;
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
@ApiModel(value="ExcelExportTemplate对象", description="excel导出模板")
public class ExcelExportTemplate implements Serializable {

    private static final long serialVersionUID=1L;
 
    @ApiModelProperty(value = "主键id", required = true)
    private Long id;
 
    @ApiModelProperty(value = "创建时间", required = true)
    @DateTimeFormat(pattern = "yyyy-MM-dd HH:mm:ss")
    @JsonFormat(timezone = "GMT+8", pattern = "yyyy-MM-dd HH:mm:ss")
    @JsonDeserialize(using = LocalDateTimeDeserializer.class)
    @JsonSerialize(using = LocalDateTimeSerializer.class)
    private LocalDateTime createTime;
 
    @ApiModelProperty(value = "创建人", required = true)
    private Long createUserId;
 
    @ApiModelProperty(value = "修改时间", required = true)
    @DateTimeFormat(pattern = "yyyy-MM-dd HH:mm:ss")
    @JsonFormat(timezone = "GMT+8", pattern = "yyyy-MM-dd HH:mm:ss")
    @JsonDeserialize(using = LocalDateTimeDeserializer.class)
    @JsonSerialize(using = LocalDateTimeSerializer.class)
    private LocalDateTime updateTime;
 
    @ApiModelProperty(value = "最后一次修改人", required = false)
    private Long updateUserId;
 
    @ApiModelProperty(value = "是否删除状态", required = true)
    private Boolean deleted;
 
    @ApiModelProperty(value = "模板名称", required = true)
    private String name;
 
    @ApiModelProperty(value = "表单配置id", required = false)
    private Long webFormTableId;
    
    @ApiModelProperty(value = "用户组id")
    private Long userGroupId;
}
