package com.framework.poi.export.excel.entity;

import java.io.Serializable;
import java.math.BigDecimal;
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
 * excel导出sheet配置
 * </p>
 *
 * @author aLiang
 * @since 2021-01-13
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
@ApiModel(value="ExcelExportSheet对象", description="excel导出sheet配置")
public class ExcelExportSheet implements Serializable {

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
 
    @ApiModelProperty(value = "sheet名字", required = false)
    private String sheetName;
 
    @ApiModelProperty(value = "标题", required = false)
    private String titleName;
 
    @ApiModelProperty(value = "描述行", required = false)
    private String description;
 
    @ApiModelProperty(value = "是否只追加，不添加标题、描述和列头", required = false)
    private Boolean onlyAppend;
 
    @ApiModelProperty(value = "间隔行数，在追加写入excel时，需要间隔多少行", required = false)
    private Integer separateNumber;
 
    @ApiModelProperty(value = "图片压缩比例", required = false)
    private BigDecimal pictureScale;
 
    @ApiModelProperty(value = "excel导出模板id", required = true)
    private Long excelExportTemplateId;
    
    @ApiModelProperty(value = "行高", required = true)
    private Integer height;
    
    @ApiModelProperty(value = "排序")
    private Integer sort;
    
    @ApiModelProperty(value = "用户组id")
    private Long userGroupId;
}
