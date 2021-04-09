package com.framework.poi.export;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFTableCell.XWPFVertAlign;

/**
 * <p>
 * 左右或上下对齐方式
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月22日 下午1:44:37
 */
public enum AlignmentEnum {

	/**
	 * 居中对齐
	 */
	CENTER(1, "居中对齐"),

	/**
	 * 左对齐或上对齐
	 */
	TOP_OR_LEFT(2, "左对齐或顶部对齐"),
	/**
	 * 右对齐或下对齐
	 */
	BOTTOM_OR_RIGHT(3, "右对齐或底部对齐");

	private Integer type;
	private String description;

	private AlignmentEnum(Integer type, String description) {
		this.type = type;
		this.description = description;
	}

	public Integer getType() {
		return type;
	}

	public String getDescription() {
		return description;
	}

	/**
	 * 
	 * <p>
	 * 段落对齐
	 * </P>
	 * 
	 * @return
	 */
	public ParagraphAlignment getParagraphAlignment() {
		if (this.equals(CENTER)) {
			return ParagraphAlignment.CENTER;
		} else if (this.equals(TOP_OR_LEFT)) {
			return ParagraphAlignment.LEFT;
		} else {
			return ParagraphAlignment.RIGHT;
		}
	}

	/**
	 * 
	 * <p>
	 * word表格垂直对齐
	 * </P>
	 * 
	 * @return
	 */
	public XWPFVertAlign getCellVerticalAlignment() {
		if (this.equals(CENTER)) {
			return XWPFVertAlign.CENTER;
		} else if (this.equals(TOP_OR_LEFT)) {
			return XWPFVertAlign.TOP;
		} else {
			return XWPFVertAlign.BOTTOM;
		}
	}
}
