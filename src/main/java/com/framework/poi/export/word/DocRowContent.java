package com.framework.poi.export.word;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Collection;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.util.CollectionUtils;

import com.framework.poi.export.AlignmentEnum;
import com.framework.poi.export.PoiUtils;
import com.framework.poi.export.Style;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;
import lombok.extern.log4j.Log4j2;

/**
 * <p>
 * 一行内容
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 上午11:33:42
 */
@Log4j2
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class DocRowContent extends DocContent {

	public static final Style DEFAULT_STYLE = new Style();
	
	/**
	 * 
	 * <p>
	 * 包含的列
	 * </P>
	 * 
	 */
	private List<? extends DocColumnContent> docColumns;

	/**
	 * 换行数量，在写入内容之前进行换行
	 */
	private Integer breakLineNumber;

	public DocRowContent(Integer contentIndex, List<? extends DocColumnContent> docColumns) {
		super(contentIndex);
		this.docColumns = docColumns;
	}

	public DocRowContent(Integer contentIndex, Integer breakLineNumber) {
		super(contentIndex);
		this.breakLineNumber = breakLineNumber;
	}
	
	public DocRowContent(Integer contentIndex) {
		super(contentIndex);
	}
	
	public DocRowContent() {
		super();
	}

	/**
	 * 
	 * <p>
	 * 添加列
	 * </P>
	 * 
	 * @param columnContent
	 * @return
	 */
	public DocRowContent addRowColumn(DocColumnContent... columnContent) {
		this.setDocColumns(this.addContent(columnContent));
		return this;
	}

	/**
	 * 
	 * <p>
	 * 处理换行
	 * </P>
	 * 
	 * @param document
	 */
	protected void handleBreakRow(XWPFDocument document) {
		Integer breakLineNumber = getBreakLineNumber();
		if (breakLineNumber != null) {
			while (breakLineNumber > 0) {
				// 换行
				document.createParagraph();
				breakLineNumber--;
			}
		}
	}

	/**
	 * 
	 * <p>
	 * 处理普通段落行
	 * </P>
	 * 
	 * @param document
	 * @param obj
	 * @param jsonAble
	 */
	public void handleRow(XWPFDocument document, Object obj, boolean jsonAble) {
		handleBreakRow(document);
		if (CollectionUtils.isEmpty(docColumns)) {
			return;
		}

		for (DocColumnContent docColumnContent : docColumns) {
			if (docColumnContent == null) {
				continue;
			}
			AlignmentEnum paragraphAlignment = docColumnContent.getParagraphAlignment();
			XWPFParagraph paragraph = document.createParagraph();
			
			Integer titleLevel = docColumnContent.getTitleLevel();
			if (titleLevel != null && titleLevel >= 0) {
				// 设置标题
				paragraph.setStyle(PoiDocExport.TITLE_SIGN + titleLevel);
			}
			
			paragraph(paragraph, paragraphAlignment);
			List<DocTextContent> docTexts = docColumnContent.getDocTexts();
			if (docTexts == null) {
				continue;
			}
			for (int i = 0; i < docTexts.size(); i++) {
				DocTextContent docTextContent = docTexts.get(i);
				if (docTextContent == null) {
					continue;
				}
				if (i > 0 && Boolean.TRUE.equals(docTextContent.getNewParagraph())) {
					paragraph = document.createParagraph();
					paragraph(paragraph, paragraphAlignment);
				}
				handleText(paragraph, docTextContent, obj, jsonAble);
			}
		}
	}

	/**
	 * 
	 * <p>
	 * 换行
	 * </P>
	 * 
	 * @param paragraph
	 */
	protected void breakRow(XWPFParagraph paragraph) {
		paragraph.createRun().addBreak();
	}

	/**
	 * 
	 * <p>
	 * 段落中写入数据
	 * </P>
	 * 
	 * @param paragraph
	 * @param docTextContent
	 * @param obj
	 * @param jsonAble
	 */
	protected void handleText(XWPFParagraph paragraph, DocTextContent docTextContent, Object obj, boolean jsonAble) {
		Boolean fixedAble = docTextContent.getFixedAble();
		// 数据值
		Object value = null;
		if (Boolean.TRUE.equals(fixedAble)) {
			// 固定值
			value = docTextContent.getFixedContent();
		} else {
			// 对象中取值
			value = PoiUtils.getValue(obj, docTextContent.getFieldName(), jsonAble);
		}
		
		if (Boolean.TRUE.equals(docTextContent.getPictureAble())) {
			// 如果是图片
			paragraphRunImages(paragraph, docTextContent.getStyle(), value, docTextContent.getPictureMonopoly());
			return;
		}

		// 前缀
		Object prefixContent = null;
		if (Boolean.FALSE.equals(docTextContent.getNoValuePrefixAble())) {
			// 空值时不需要前缀
			prefixContent = value == null ? null : docTextContent.getPrefixContent();
		} else {
			prefixContent = docTextContent.getPrefixContent();
		}

		Object suffixContent = null;
		if (Boolean.FALSE.equals(docTextContent.getNoValueSuffixAble())) {
			// 空值时不需要后缀
			suffixContent = value == null ? null : docTextContent.getSuffixContent();
		} else {
			suffixContent = docTextContent.getSuffixContent();
		}

		if (value == null) {
			// 如果值是空的则使用默认值
			value = docTextContent.getDefaultContent();
		}

		// 值转换
		value = PoiUtils.formatValue(value, docTextContent.getDateTimeFromat(), docTextContent.getValueExchange(), jsonAble);
		prefixContent = valueValidate(prefixContent);
		suffixContent = valueValidate(suffixContent);
		// 需要写入word中的内容
		String content = prefixContent.toString() + value.toString() + suffixContent.toString();
		// 写入内容
		paragraphRun(paragraph, docTextContent.getStyle(), content);
	}

	protected Object valueValidate(Object obj) {
		if (obj == null) {
			return "";
		}
		return obj;
	}
	
	/**
	 * 
	 * <p>
	 * 段落信息
	 * </P>
	 * @param paragraph
	 * @param paragraphAlignment
	 */
	protected void paragraph(XWPFParagraph paragraph, AlignmentEnum paragraphAlignment) {
		if (paragraphAlignment == null) {
			// 默认左对齐
			paragraphAlignment = AlignmentEnum.TOP_OR_LEFT;
		}
		// 设置对齐方式
		paragraph.setAlignment(paragraphAlignment.getParagraphAlignment());
	}

	/**
	 * 
	 * <p>
	 * 段落中的内容
	 * </P>
	 * 
	 * @param paragraph
	 * @param style
	 * @param value
	 */
	protected void paragraphRun(XWPFParagraph paragraph, Style style, String value) {
		if (StringUtils.isEmpty(value)) {
			return;
		}
		if (style == null) {
			style = DEFAULT_STYLE;
		}
		Integer fontSize = style.getFontSize();
		if (fontSize == null) {
			fontSize = Style.DEFAULT_FONT_SIZE;
		}

		String fontType = style.getFontType();
		if (fontType == null) {
			fontType = Style.DEFAULT_FONT_TYPE;
		}
		// 一个XWPFRun代表具有相同属性的一个区域：一段文本
		XWPFRun run = paragraph.createRun();
		// 设置字体
		run.setFontFamily(fontType);
		// 设置字体大小
		run.setFontSize(fontSize);

		String fontColor = style.getFontColor();
		if (!StringUtils.isEmpty(fontColor)) {
			run.setColor(fontColor);
		}
		run.setBold(Boolean.TRUE.equals(style.getBoldAble()));
		run.setText(value);
	}

	/**
	 * 
	 * <p>
	 * 导出图片
	 * </P>
	 * @param paragraph
	 * @param style
	 * @param value
	 * @param pictureMonopoly
	 */
	private void paragraphRunImages(XWPFParagraph paragraph, Style style, Object value, Boolean pictureMonopoly) {
		if (value == null) {
			return;
		}
		if (value instanceof Collection) {
			if (pictureMonopoly == null) {
				// 多张图片默认换行展示
				pictureMonopoly = false;
			}
			Collection<?> images = (Collection<?>) value;
			int i = 0;
			for (Object object : images) {
				paragraphRunImage(paragraph, style, object);
				if (!pictureMonopoly && i != images.size() - 1) {
					// 换行展示，最后一行不换行，否则会多出一行换行符
					breakRow(paragraph);
				}
				i++;
			}
		} else {
			paragraphRunImage(paragraph, style, value);
		}
	}

	/**
	 * 
	 * <p>
	 * 导出图片
	 * </P>
	 * 
	 * @param paragraph
	 * @param style
	 * @param value
	 */
	private void paragraphRunImage(XWPFParagraph paragraph, Style style, Object value) {
		if (value == null) {
			return;
		}
		Path path = Paths.get(value.toString());
		InputStream stream = null;
		try {
			if (Files.notExists(path)) {
				log.error("图片文件不存在{}", value);
				return;
			}
			Integer width = style.getWidth();
			if (width == null) {
				width = 160;
			}
			Integer height = style.getHeight();
			if (height == null) {
				height = 160;
			}
			stream = Files.newInputStream(path);
			XWPFRun imageCellRunn = paragraph.createRun();
			imageCellRunn.addPicture(stream, XWPFDocument.PICTURE_TYPE_JPEG, "", Units.toEMU(width),
					Units.toEMU(height));
		} catch (Exception e) {
			log.error("导出图片失败，图片路径:{}", value, e);
		} finally {
			if (stream != null) {
				try {
					stream.close();
				} catch (IOException e) {
					log.error("导出图片关闭流失败，图片路径:{}", value, e);
				}
			}
		}

	}

}
