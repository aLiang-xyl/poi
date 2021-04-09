package com.framework.poi.export.word;

import com.framework.poi.export.Style;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * <p>
 * 文本内容
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月19日 下午2:07:30
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class DocTextContent extends DocContent {

	/**
	 * 是否是固定内容
	 */
	private Boolean fixedAble;

	/**
	 * 是否是标题，如果是标题，将会成为目录
	 */
	private Boolean titleAble;

	/**
	 * 固定内容信息
	 */
	private String fixedContent;

	/**
	 * 字段名，如果是从对象中取值，则使用该字段
	 */
	private String fieldName;

	/**
	 * 后缀内容
	 */
	private String suffixContent;

	/**
	 * 没有值时是否追加后缀内容
	 */
	private Boolean noValueSuffixAble;

	/**
	 * 前缀内容，会追加在固定内容或者取值内容前面
	 */
	private String prefixContent;

	/**
	 * 没有值时是否追加前缀内容
	 */
	private Boolean noValuePrefixAble;

	/**
	 * 是否使用新的段落来写入（使用新的行）
	 */
	private Boolean newParagraph;

	/**
	 * 默认值
	 */
	private String defaultContent;

	/**
	 * 日期格式，日期对象必须是LocalTime、LocalDate、LocalDateTime
	 * 
	 * 
	 */
	private String dateTimeFromat;

	/**
	 * 是否是图片<br/>
	 * 如果是一张图片，可以使用String类型值为图片的绝对路径<br/>
	 * 如果是多张图，必须是List&lt; String &gt;类型，元素是图片的绝对路径
	 * 
	 */
	private Boolean pictureAble;

	/**
	 * 图片是否独占一行<br>
	 * 如果图片不是独占一行，且是多张图片的时候,图片会换行展示
	 */
	private Boolean pictureMonopoly;

	/**
	 * 值转换，valueExchange={"1":"男", "0":"女"},当对应的变量值为1时则转换成男
	 * 
	 * 
	 */
	private String valueExchange;
	
	/**
	 * 样式信息
	 */
	private Style style;

	public DocTextContent() {
		super();
	}

	public DocTextContent(Integer contentIndex) {
		super(contentIndex);
	}

	public DocTextContent(Integer contentIndex, String fieldName, Style style) {
		super(contentIndex);
		this.fieldName = fieldName;
		this.style = style;
	}

	public DocTextContent(Integer contentIndex, Boolean fixedAble, String fixedContent, Style style) {
		super(contentIndex);
		this.fixedAble = fixedAble;
		this.fixedContent = fixedContent;
		this.style = style;
	}

}
