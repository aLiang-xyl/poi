package com.framework.poi.export.word.table;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.springframework.util.CollectionUtils;

import com.framework.poi.export.AlignmentEnum;
import com.framework.poi.export.ColumnMerge;
import com.framework.poi.export.ColumnStyle;
import com.framework.poi.export.PoiUtils;
import com.framework.poi.export.word.DocColumnContent;
import com.framework.poi.export.word.DocRowContent;
import com.framework.poi.export.word.DocTextContent;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * <p>
 * 表格行信息
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 下午5:45:50
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class DocTableRowContent extends DocRowContent {

	private static final ColumnStyle CELL_STYLE = new ColumnStyle();
	
	/**
	 * 合并信息
	 */
	private List<ColumnMerge> merges;

	/**
	 * 字段名，如果是从对象中取值，则使用该字段，如果取值的结果是一个List集合，则会循环该行
	 */
	private String fieldName;
	
	/**
	 * 行高
	 */
	private Integer height;
	
	private XWPFTable table;

	public DocTableRowContent(Integer contentIndex, List<DocTableColumnContent> docColumns) {
		super(contentIndex, docColumns);
	}

	public DocTableRowContent(Integer contentIndex, Integer height) {
		super(contentIndex);
		this.height = height;
	}
	
	public DocTableRowContent(Integer contentIndex) {
		super(contentIndex);
	}

	public DocTableRowContent() {
		super();
	}

	@Override
	public void handleRow(XWPFDocument document, Object obj, boolean jsonAble) {
		handleBreakRow(document);
		if (table == null) {
			table = document.createTable();
			// 默认有一行，需要删掉
			table.removeRow(0);
		}
//		setTableWidth(table);
		hanleTableRow(table, obj, jsonAble);
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
	public DocTableRowContent addRowColumn(DocTableColumnContent... columnContent) {
		this.setDocColumns(this.addContent(columnContent));
		return this;
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
	public DocTableRowContent addMerge(ColumnMerge merge) {
		if (this.merges == null) {
			this.merges = new ArrayList<>();
		}
		this.merges.add(merge);
		return this;
	}

	/**
	 * 
	 * <p>
	 * 处理表格行
	 * </P>
	 * 
	 * @param table
	 * @param obj
	 * @param jsonAble
	 */
	private void hanleTableRow(XWPFTable table, Object obj, boolean jsonAble) {
		List<? extends DocColumnContent> docColumns = this.getDocColumns();
		if (CollectionUtils.isEmpty(docColumns)) {
			return;
		}
		String fieldName = this.getFieldName();
		if (!StringUtils.isEmpty(fieldName)) {
			obj = PoiUtils.getValue(obj, fieldName, jsonAble);
		}
		if (obj != null && obj instanceof List) {
			List<?> list = (List<?>) obj;
			for (Object object : list) {
				createTableRow(table, object, jsonAble);
			}
		} else {
			XWPFTableRow tableRow = createTableRow(table, obj, jsonAble);
			handleColumnMerge(tableRow);
		}
	}

	/**
	 * 
	 * <p>
	 * 跨列合并
	 * </P>
	 * 
	 * @param tableRow
	 */
	private void handleColumnMerge(XWPFTableRow tableRow) {
		if (this.merges == null) {
			return;
		}
		for (ColumnMerge merge : merges) {
			Integer start = merge.getMergeColumnStart();
			Integer end = merge.getMergeColumnEnd();
			if (start == null || end == null || end <= start) {
				return;
			}
			for (int i = 0; i <= end; i++) {
				XWPFTableCell cell = tableRow.getCell(i);
				if (cell == null) {
					cell = tableRow.createCell();
				}
				if (i == start || i == end + 1) {
					cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
				} else {
					cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
				}
			}
		}
	}

	public void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
		for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
			XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
			if (rowIndex == fromRow) {
				// The first merged cell is set with RESTART merge value
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
			} else {
				// Cells which join (merge) the first one,are set with CONTINUE
				cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
			}
		}
	}

	/**
	 * 
	 * <p>
	 * 创建表格行信息
	 * </P>
	 * 
	 * @param table
	 * @param obj
	 * @param jsonAble
	 */
	@SuppressWarnings("unchecked")
	private XWPFTableRow createTableRow(XWPFTable table, Object obj, boolean jsonAble) {
		List<DocTableColumnContent> docColumns = (List<DocTableColumnContent>) this.getDocColumns();
		if (CollectionUtils.isEmpty(docColumns)) {
			return null;
		}
		XWPFTableRow tableRow = table.createRow();
		if (this.height != null) {
			tableRow.setHeight(this.height);
		}

		for (DocTableColumnContent docColumnContent : docColumns) {
			ColumnStyle style = docColumnContent.getStyle();
			if (style == null) {
				style = CELL_STYLE;
			}
			AlignmentEnum verticalAlignment = style.getVerticalAlignment();
			if (verticalAlignment == null) {
				verticalAlignment = AlignmentEnum.CENTER;
			}
			Integer contentIndex = docColumnContent.getContentIndex();
			// 表格信息
			XWPFTableCell tableCell = tableRow.getCell(contentIndex);
			while (tableCell == null) {
				tableRow.createCell();
				tableCell = tableRow.getCell(contentIndex);
			}
			// 默认有一个段落，需要删掉，不然是一个换行
			tableCell.removeParagraph(0);
			// 表格垂直对齐方式， 表格没有左右对齐，左右对齐是段落的对齐
			tableCell.setVerticalAlignment(verticalAlignment.getCellVerticalAlignment());
			tableCell.setColor(style.getBackgroundColor());
			if (style.getWidth() != null) {
//				setCellWidth(table, style.getWidth());
				setCellWidth(tableCell, style.getWidth());
			}
			// 处理表格列信息
			handleTableColumn(tableCell, docColumnContent, obj, jsonAble);
		}
		return tableRow;
	}

	/**
	 * 
	 * <p>
	 * 设置单元格宽度
	 * </P>
	 * 
	 * @param table
	 * @param width
	 */
//	private void setCellWidth(XWPFTable table, Integer width) {
//		if (width == null) {
//			return;
//		}
//		CTTbl ttbl = table.getCTTbl();
//		if (ttbl.getTblGrid() != null) {
//			return;
//		}
//		CTTblGrid tblGrid = ttbl.addNewTblGrid();
//		CTTblGridCol gridCol = tblGrid.addNewGridCol();
//		gridCol.setW(new BigInteger(width.toString()));
//		System.out.println(gridCol.isSetW());
//		CTTblWidth tblWidth = cellPr.isSetTcW() ? cellPr.getTcW() : cellPr.addNewTcW();
//		if(!StringUtils.isEmpty(width)){
//			tblWidth.setW(new BigInteger(width));
//			tblWidth.setType(STTblWidth.DXA);
//		}
//
//	}

	/**
	 * 
	 * <p>
	 * 列设置宽度
	 * </P>
	 * 
	 * @param cell
	 * @param width
	 */
	private void setCellWidth(XWPFTableCell cell, Integer width) {
		if (width == null) {
			return;
		}
		CTTcPr cellPr = getCellCTTcPr(cell);
		CTTblWidth tblWidth = cellPr.isSetTcW() ? cellPr.getTcW() : cellPr.addNewTcW();
		tblWidth.setW(new BigInteger(width.toString()));
		tblWidth.setType(STTblWidth.DXA);
	}

	private CTTcPr getCellCTTcPr(XWPFTableCell cell) {
		CTTc cttc = cell.getCTTc();
		CTTcPr tcPr = cttc.isSetTcPr() ? cttc.getTcPr() : cttc.addNewTcPr();
		return tcPr;
	}

	/**
	 * 
	 * <p>
	 * 设置表格行宽度
	 * </P>
	 * 
	 * @param table
	 */
//	private void setTableWidth(XWPFTable table) {
//		Style style = this.getStyle();
//		String width = null;
//		if (style != null && style.getWidth() != null) {
//			width = style.getWidth().toString();
//		} else {
//			width = "8300";
//		}
//		CTTbl ttbl = table.getCTTbl();
//		CTTblPr tblPr = ttbl.getTblPr() == null ? ttbl.addNewTblPr() : ttbl.getTblPr();
//		CTTblWidth tblWidth = tblPr.isSetTblW() ? tblPr.getTblW() : tblPr.addNewTblW();
//		tblWidth.setW(new BigInteger(width));
//		tblWidth.setType(STTblWidth.DXA);
//	}

	/**
	 * 
	 * <p>
	 * 处理表格列信息
	 * </P>
	 * 
	 * @param tableCell
	 * @param docColumnContent
	 * @param obj
	 * @param jsonAble
	 */
	private void handleTableColumn(XWPFTableCell tableCell, DocColumnContent docColumnContent, Object obj,
			boolean jsonAble) {
		List<DocTextContent> docTexts = docColumnContent.getDocTexts();
		if (docTexts == null) {
			return;
		}
		AlignmentEnum paragraphAlignment = docColumnContent.getParagraphAlignment();
		XWPFParagraph paragraph = tableCell.addParagraph();
		paragraph(paragraph, paragraphAlignment);

		for (int i = 0; i < docTexts.size(); i++) {
			DocTextContent docTextContent = docTexts.get(i);
			if (docTextContent == null) {
				continue;
			}
			if (i > 0 && Boolean.TRUE.equals(docTextContent.getNewParagraph())) {
				paragraph = tableCell.addParagraph();
				paragraph(paragraph, paragraphAlignment);
			}
			handleText(paragraph, docTextContent, obj, jsonAble);
		}
	}
}
