package com.framework.poi.export.word;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTOnOff;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTStyle;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STStyleType;
import org.springframework.util.StringUtils;

import com.framework.poi.export.Assert;
import com.framework.poi.export.BusinessException;

import lombok.Getter;
import lombok.extern.log4j.Log4j2;

/**
 * <p>
 * word导出
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月29日 下午6:00:42
 */
@Log4j2
public class PoiDocExport {

	private static final String TITLE_NAME = "标题";
	public static final String TITLE_SIGN = "TOC";

	private XWPFDocument document;

	/**
	 * 最大标题级别
	 */
	@Getter
	private Integer maxTitleLevel;

	/**
	 * 
	 * <p>
	 * 导出实例化
	 * </P>
	 * 
	 * @param maxTitleLevel 最大标题级别
	 */
	public PoiDocExport(Integer maxTitleLevel) {
		super();
		this.document = new XWPFDocument();
		this.maxTitleLevel = maxTitleLevel;
		createTitle();
	}

	public PoiDocExport() {
		this(null);
	}

	public void export(DocAreaContent areaContent) {
		areaContent.export(document);
	}

	/**
	 * 
	 * <p>
	 * 写入
	 * </P>
	 * 
	 * @param excelPath 文件路径
	 */
	public File write(String excelPath) {
		Assert.isTrue(!StringUtils.isEmpty(excelPath), "路径参数错误");
		OutputStream newOutputStream = null;
		Path path = Paths.get(excelPath);
		try {
			newOutputStream = Files.newOutputStream(Paths.get(excelPath));
			write(newOutputStream);
			return path.toFile();
		} catch (IOException e) {
			log.error("导出文件失败", e);
			throw new BusinessException("导出文件失败");
		} finally {
			if (newOutputStream != null) {
				try {
					newOutputStream.close();
				} catch (IOException e) {
					log.error("导出文件关闭流失败", e);
				}
			}
		}

	}

	/**
	 * 
	 * <p>
	 * 写入文件
	 * </P>
	 * 
	 * @param outputStream
	 */
	public void write(OutputStream outputStream) {
		try {
			this.document.write(outputStream);
		} catch (IOException e) {
			log.error("导出一行出错", e);
		}
	}

	/**
	 * 
	 * <p>
	 * 创建标题
	 * </P>
	 */
	private void createTitle() {
		if (this.maxTitleLevel == null || this.maxTitleLevel <= 0) {
			return;
		}
		for (int i = 1; i <= this.maxTitleLevel; i++) {
			addCustomHeadingStyle(TITLE_NAME + i, TITLE_SIGN + i, i);
		}
	}

	/**
	 * 
	 * <p>
	 * 添加标题
	 * </P>
	 * 
	 * @param name
	 * @param strStyleId
	 * @param headingLevel
	 */
	private void addCustomHeadingStyle(String name, String strStyleId, int headingLevel) {

		CTStyle ctStyle = CTStyle.Factory.newInstance();
		ctStyle.setStyleId(strStyleId);

		CTString styleName = CTString.Factory.newInstance();
		styleName.setVal(name);
		ctStyle.setName(styleName);

		CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
		indentNumber.setVal(BigInteger.valueOf(headingLevel));

		// lower number > style is more prominent in the formats bar
		ctStyle.setUiPriority(indentNumber);

		CTOnOff onoffnull = CTOnOff.Factory.newInstance();
		ctStyle.setUnhideWhenUsed(onoffnull);

		// style shows up in the formats bar
		ctStyle.setQFormat(onoffnull);

		// style defines a heading of the given level
		CTPPr ppr = CTPPr.Factory.newInstance();
		ppr.setOutlineLvl(indentNumber);
		ctStyle.setPPr(ppr);

		XWPFStyle style = new XWPFStyle(ctStyle);

		// is a null op if already defined
		XWPFStyles styles = this.document.createStyles();

		style.setType(STStyleType.PARAGRAPH);
		styles.addStyle(style);

	}
}
