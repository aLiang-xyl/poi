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
 * excel表格列配置
 * </p>
 *
 * @author aLiang
 * @since 2021-01-13
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
@ApiModel(value = "ExcelColumn对象", description = "excel表格列配置")
public class ExcelColumn implements Serializable {

	private static final long serialVersionUID = 1L;

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
	
	@ApiModelProperty(value = "用户组id")
    private Long userGroupId;

	/**
	 * 列序号，根据序号会从第一列依次填充，从0开始
	 * 
	 * 
	 */
	@ApiModelProperty(value = "列序号，根据序号会从第一列依次填充，从0开始", required = true)
	private Integer columnIndex;

	/**
	 * 列标题名称
	 * 
	 */
	@ApiModelProperty(value = "列标题名称", required = true)
	private String titleName;

	/**
	 * 列名，用来取值，多层级取值示例：person.name, company.person.name
	 * 
	 */
	@ApiModelProperty(value = "列名，用来取值，多层级取值示例：person.name, company.person.name", required = true)
	private String fieldName;

	/**
	 * 追加信息
	 */
	@ApiModelProperty(value = "追加信息", required = false)
	private String append;

	/**
	 * 列宽，默认-1则使用默认列宽度
	 * 
	 * 
	 */
	@ApiModelProperty(value = "列宽", required = true)
	private Integer width;

	/**
	 * 日期格式，日期对象必须是LocalTime、LocalDate、LocalDateTime
	 * 
	 * 
	 */
	@ApiModelProperty(value = "日期格式", required = false)
	private String dateTimeFromat;

	/**
	 * 是否是图片<br/>
	 * 如果是一张图片，可以使用String类型值为图片的绝对路径<br/>
	 * 如果是多张图，必须是List&lt; String &gt;类型，元素是图片的绝对路径
	 * 
	 */
	@ApiModelProperty(value = "是否是图片", required = false)
	private Boolean pictureAble;

	/**
	 * 图片是否独占一格，如果图片独占一格，且是多张图片的时候需要放在最后一列
	 */
	@ApiModelProperty(value = "图片是否独占一格", required = false)
	private Boolean pictureMonopoly;

	/**
	 * 图片高度百分比，相对于该行的行高，默认紧靠边框
	 * 
	 * 
	 */
	@ApiModelProperty(value = "图片高度百分比", required = false)
	private BigDecimal picHeight;

	/**
	 * 图片宽度百分比，相对于该列的宽度，默认紧靠边框
	 * 
	 * 
	 */
	@ApiModelProperty(value = "图片高度百分比", required = false)
	private BigDecimal picWidth;

	/**
	 * 值转换，valueExchange={"1":"男", "0":"女"},当对应的变量值为1时则转换成男
	 * 
	 * 
	 */
	@ApiModelProperty(value = "值转换", required = false)
	private String valueExchange;

	/**
	 * 
	 * <p>
	 * 合并时依赖的字段，多层级取值示例：person.name, company.person.name
	 * </P>
	 * 
	 * 
	 */
	@ApiModelProperty(value = "合并时依赖的字段，多层级取值示例：person.name, company.person.name", required = false)
	private String mergeKeyField;

	/**
	 * 
	 * <p>
	 * 是否合并单元格，如果需要合并单元格，需要先按照合并规则排序，否则合并会出现问题
	 * </P>
	 * 
	 * 
	 */
	@ApiModelProperty(value = "是否合并单元格", required = false)
	private Boolean mergeAble;
	
	@ApiModelProperty(value = "默认值", required = false)
	private String defaultValue;

	@ApiModelProperty(value = "excel导出sheet id", required = true)
	private Long excelExportSheetId;
	
	public ExcelColumn(Integer columnIndex, String titleName, String fieldName) {
		super();
		this.columnIndex = columnIndex;
		this.titleName = titleName;
		this.fieldName = fieldName;
	}

	public ExcelColumn() {
		super();
	}
}
