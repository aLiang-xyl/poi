package com.framework.poi.export.word.table;

import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.springframework.util.CollectionUtils;

import com.framework.poi.export.Assert;
import com.framework.poi.export.RowMerge;
import com.framework.poi.export.word.DocAreaContent;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;
import lombok.extern.log4j.Log4j2;

/**
 * <p>
 * 域内容
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 下午5:32:19
 */
@Log4j2
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class DocTableAreaContent extends DocAreaContent {

	/**
	 * 行合并
	 */
	private List<RowMerge> merges;

	/**
	 * 
	 * <p>
	 * 导出
	 * </P>
	 * 
	 * @param document
	 * @see com.framework.poi.export.word.DocAreaContent#export(org.apache.poi.xwpf.usermodel.XWPFDocument)
	 */
	@SuppressWarnings("unchecked")
	public void export(XWPFDocument document) {
		List<DocTableRowContent> docRows = (List<DocTableRowContent>) this.getDocRows();
		Assert.isTrue(!CollectionUtils.isEmpty(docRows), "参数错误");
		XWPFTable table = document.createTable();
		// 默认有一行，需要删掉
		table.removeRow(0);
		for (DocTableRowContent docRowContent : docRows) {
			if (docRowContent == null) {
				continue;
			}
			docRowContent.setTable(table);
			docRowContent.handleRow(document, super.getObj(), super.isJsonAble());
		}

		// 处理行合并
		handleRowMerge(table, merges);
	}

	public DocTableAreaContent(List<DocTableRowContent> docRows, Object obj, boolean jsonAble) {
		super(null, obj, jsonAble);
	}

	public DocTableAreaContent() {
		super();
	}

	/**
	 * 
	 * <p>
	 * 处理行合并
	 * </P>
	 * 
	 * @param table
	 * @param docRowContent
	 * @param col
	 */
	private void handleRowMerge(XWPFTable table, List<RowMerge> merges) {
		if (CollectionUtils.isEmpty(merges)) {
			return;
		}
		for (RowMerge merge : merges) {
			Integer mergeRowStart = merge.getMergeRowStart();
			Integer mergeRowEnd = merge.getMergeRowEnd();
			if (mergeRowStart == null || mergeRowEnd == null) {
				continue;
			}
			for (int rowIndex = mergeRowStart; rowIndex <= mergeRowEnd; rowIndex++) {
				XWPFTableRow row = table.getRow(rowIndex);
				if (row == null) {
					log.error("行合并时出错，不存在对应的行，行索引:{}", rowIndex);
					continue;
				}
				
				Integer columnIndex = merge.getColumnIndex();
				XWPFTableCell cell = row.getCell(columnIndex);
				if (cell == null) {
					log.error("行合并时出错，不存在对应的列，列索引:{}", columnIndex);
					continue;
				}
				if (rowIndex == mergeRowStart) {
					// The first merged cell is set with RESTART merge value
					getCellCTTcPr(cell).addNewVMerge().setVal(STMerge.RESTART);
				} else {
					// Cells which join (merge) the first one,are set with CONTINUE
					getCellCTTcPr(cell).addNewVMerge().setVal(STMerge.CONTINUE);
				}
			}
		}
	}

	private CTTcPr getCellCTTcPr(XWPFTableCell cell) {
		CTTc cttc = cell.getCTTc();
		CTTcPr tcPr = cttc.isSetTcPr() ? cttc.getTcPr() : cttc.addNewTcPr();
		return tcPr;
	}
}
