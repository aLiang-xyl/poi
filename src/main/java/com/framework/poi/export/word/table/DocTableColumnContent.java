package com.framework.poi.export.word.table;

import java.util.ArrayList;
import java.util.List;

import com.framework.poi.export.AlignmentEnum;
import com.framework.poi.export.ColumnStyle;
import com.framework.poi.export.RowMerge;
import com.framework.poi.export.word.DocColumnContent;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;
import lombok.experimental.Accessors;

/**
 * <p>
 * 表格列信息
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 下午6:17:11
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
@AllArgsConstructor
@NoArgsConstructor
public class DocTableColumnContent extends DocColumnContent {

	/**
	 * 表格样式
	 */
	private ColumnStyle style;
	
	/**
	 * 合并信息
	 */
	private List<RowMerge> merges;

	public DocTableColumnContent(Integer contentIndex) {
		super(contentIndex);
	}
	
	public DocTableColumnContent(Integer contentIndex, AlignmentEnum paragraphAlignment) {
		super(contentIndex, paragraphAlignment);
	}

	public DocTableColumnContent(Integer contentIndex, AlignmentEnum paragraphAlignment, ColumnStyle style) {
		super(contentIndex, paragraphAlignment);
		this.style = style;
	}
	
	public DocTableColumnContent(Integer contentIndex, ColumnStyle style) {
		super(contentIndex);
		this.style = style;
	}
	
	public DocTableColumnContent(Integer contentIndex, ColumnStyle style, List<RowMerge> merges) {
		super(contentIndex);
		this.style = style;
		this.merges = merges;
	}
	
	/**
	 * 
	 * <p>
	 * 居中展示列
	 * </P>
	 * @param contentIndex
	 * @param style
	 * @return
	 */
	public static DocTableColumnContent centerTableColumn(Integer contentIndex, ColumnStyle style) {
		return new DocTableColumnContent(contentIndex, AlignmentEnum.CENTER, style);
	}

	/**
	 * 
	 * <p>
	 * 添加合并信息
	 * </P>
	 * 
	 * @param merge
	 * @return
	 */
	public DocTableColumnContent addMerge(RowMerge merge) {
		if (this.merges == null) {
			this.merges = new ArrayList<>();
		}
		this.merges.add(merge);
		return this;
	}
}
