package com.framework.poi.export.word;

import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.util.CollectionUtils;

import com.framework.poi.export.Assert;

import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.experimental.Accessors;

/**
 * <p>
 * 域内容
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月21日 下午5:32:19
 */
@Data
@EqualsAndHashCode(callSuper = false)
@Accessors(chain = true)
public class DocAreaContent {

	/**
	 * 行内容
	 */
	private List<? extends DocRowContent> docRows;

	/**
	 * 对象信息
	 */
	private Object obj;

	/**
	 * 是否是json
	 */
	private boolean jsonAble;

	/**
	 * 
	 * <p>
	 * 导出
	 * </P>
	 * 
	 * @param document
	 */
	public void export(XWPFDocument document) {
		List<? extends DocRowContent> docRows = this.getDocRows();
		Assert.isTrue(!CollectionUtils.isEmpty(docRows), "参数错误");
		for (DocRowContent docRowContent : docRows) {
			if (docRowContent == null) {
				continue;
			}
			docRowContent.handleRow(document, obj, jsonAble);
		}
	}

	public DocAreaContent(List<DocRowContent> docRows, Object obj, boolean jsonAble) {
		super();
		this.docRows = docRows;
		this.obj = obj;
		this.jsonAble = jsonAble;
	}

	public DocAreaContent() {
		super();
	}

}
