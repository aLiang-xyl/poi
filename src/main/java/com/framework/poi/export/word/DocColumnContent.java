package com.framework.poi.export.word;

import java.util.LinkedList;
import java.util.List;

import com.framework.poi.export.AlignmentEnum;
import com.framework.poi.export.Style;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * <p>
 * 列内容
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 下午3:21:54
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class DocColumnContent extends DocContent {

	/**
	 * 段落对齐方式，左、右、居中
	 */
	private AlignmentEnum paragraphAlignment; 
	
	/**
	 * 标题级别
	 */
	private Integer titleLevel;
	
	/**
	 * 
	 * <p>
	 * 文本配置信息
	 * </P>
	 * 
	 * @return
	 */
	private List<DocTextContent> docTexts;
	
	public DocColumnContent() {
		super();
	}

	public DocColumnContent(Integer contentIndex) {
		super(contentIndex);
	}
	
	public DocColumnContent(Integer contentIndex, AlignmentEnum paragraphAlignment) {
		super(contentIndex);
		this.paragraphAlignment = paragraphAlignment;
	}

	public DocColumnContent(Integer contentIndex, List<DocTextContent> docTexts) {
		super(contentIndex);
		this.docTexts = docTexts;
	}
	
	/**
	 * 
	 * <p>
	 * 居中展示列
	 * </P>
	 * @param contentIndex
	 * @return
	 */
	public static DocColumnContent centerColumn(Integer contentIndex) {
		return new DocColumnContent(contentIndex, AlignmentEnum.CENTER);
	}

	/**
	 * 
	 * <p>
	 * 添加文本元素
	 * </P>
	 * 
	 * @param docTextContent
	 */
	public DocColumnContent addDocText(DocTextContent... docTextContent) {
		this.setDocTexts(this.addContent(docTextContent));
		return this;
	}

	/**
	 * 
	 * <p>
	 * 添加文本元素
	 * </P>
	 * 
	 * @param textContentIndex
	 * @param fieldName
	 * @param style
	 * @return
	 */
	public DocColumnContent addDocText(Integer textContentIndex, String fieldName, Style style) {
		if (this.docTexts == null) {
			docTexts = new LinkedList<DocTextContent>();
		}
		docTexts.add(new DocTextContent(textContentIndex, fieldName, style));
		return this;
	}

	/**
	 * 
	 * <p>
	 * 添加文本元素
	 * </P>
	 * 
	 * @param textContentIndex
	 * @param fixedAble
	 * @param fixedContent
	 * @param style
	 * @return
	 */
	public DocColumnContent addDocText(Integer textContentIndex, Boolean fixedAble, String fixedContent, Style style) {
		if (this.docTexts == null) {
			docTexts = new LinkedList<DocTextContent>();
		}
		docTexts.add(new DocTextContent(textContentIndex, fixedAble, fixedContent, style));
		return this;
	}
}
