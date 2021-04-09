package com.framework.poi.export.word.header;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHdrFtr;
import org.springframework.util.CollectionUtils;

import com.framework.poi.export.StreamUtils;
import com.framework.poi.export.word.DocColumnContent;
import com.framework.poi.export.word.DocRowContent;
import com.framework.poi.export.word.DocTextContent;

/**
 * 
 * <p>
 * 页眉行信息
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月28日 下午4:20:55
 */
public class DocHeaderRowContent extends DocRowContent {

	@Override
	public void handleRow(XWPFDocument document, Object obj, boolean jsonAble) {
		List<? extends DocColumnContent> docColumns = super.getDocColumns();
		if (CollectionUtils.isEmpty(docColumns)) {
			return;
		}
		// 创建页眉
		CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
		XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy( document, sectPr );
		XWPFHeader header = headerFooterPolicy.createHeader( STHdrFtr.DEFAULT );
		
		List<XWPFParagraph> list = new ArrayList<>();
		for (DocColumnContent docColumnContent : docColumns) {
			if (CollectionUtils.isEmpty(docColumnContent.getDocTexts())) {
				continue;
			}
			XWPFParagraph paragraph = header.createParagraph();

			super.paragraph(paragraph, docColumnContent.getParagraphAlignment());
			handleColumn(docColumnContent, paragraph, obj, jsonAble);
			list.add(paragraph);
		}
		// 处理页眉图片关系
		Map<Long, XWPFPictureData> map = StreamUtils.groupingByGetOne(header.getAllPackagePictures(), XWPFPictureData::getChecksum);
		for (XWPFParagraph xwpfParagraph : list) {
			List<XWPFRun> runs = xwpfParagraph.getRuns();
			for (XWPFRun run : runs) {
				List<XWPFPicture> embeddedPictures = run.getEmbeddedPictures();
				for (XWPFPicture picture : embeddedPictures) {
					Long checksum = picture.getPictureData().getChecksum();
					XWPFPictureData xwpfPictureData = map.get(checksum);
					if (xwpfPictureData == null) {
						continue;
					}
					picture.getCTPicture().getBlipFill().getBlip().setEmbed(header.getRelationId(xwpfPictureData));
				}
			}
		}
	}

	/**
	 * 
	 * <p>
	 * 处理列
	 * </P>
	 * 
	 * @param columnContent
	 * @param paragraph
	 * @param obj
	 * @param jsonAble
	 */
	private void handleColumn(DocColumnContent columnContent, XWPFParagraph paragraph, Object obj, boolean jsonAble) {
		List<DocTextContent> docTexts = columnContent.getDocTexts();
		for (DocTextContent docTextContent : docTexts) {
			super.handleText(paragraph, docTextContent, obj, jsonAble);
		}
	}
}
