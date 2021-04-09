package com.framework.common;

import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;
import java.util.UUID;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.framework.poi.export.AlignmentEnum;
import com.framework.poi.export.ColumnMerge;
import com.framework.poi.export.ColumnStyle;
import com.framework.poi.export.RowMerge;
import com.framework.poi.export.Style;
import com.framework.poi.export.word.DocAreaContent;
import com.framework.poi.export.word.DocColumnContent;
import com.framework.poi.export.word.DocRowContent;
import com.framework.poi.export.word.DocTextContent;
import com.framework.poi.export.word.PoiDocExport;
import com.framework.poi.export.word.header.DocHeaderRowContent;
import com.framework.poi.export.word.table.DocTableAreaContent;
import com.framework.poi.export.word.table.DocTableColumnContent;
import com.framework.poi.export.word.table.DocTableRowContent;

public class WordTemplateTest {

	private static String color = "F5F5F5";
	private static Style titleStyle = getTitleStyle().setBoldAble(true);
	private static Style minTitleStyle = getMinTitleStyle().setBoldAble(true);
	// 加黑正文
	private static Style boldTextStyle = getTextStyle().setBoldAble(true);
	// 正文
	private static Style textStyle = getTextStyle();
	
	private static Integer DEFAULT_HEIGHT = 600;
	// 表格样式
	private static ColumnStyle tableCenterStyle = ColumnStyle.centerStyle(3300);
//	// 表格样式
//	private static Style tableCenterStyle = Style.tableCenterStyle(3300);

	public static void main(String[] args) {
		DocAreaContent areaContent = new DocAreaContent();

		areaContent.setDocRows(test());

		JSONObject obj = new JSONObject();
		obj.put("name", "xxx集团");
		obj.put("biaoduan", "xxx项目");
		obj.put("code", "12345678");
		obj.put("test1", "这是测试数据");
		obj.put("mianji", "3.78");
		areaContent.setObj(obj);
		areaContent.setJsonAble(true);

		List<DocRowContent> baseDataTable = getBaseDataTable();
		DocTableAreaContent baseTableContent = new DocTableAreaContent();
		baseTableContent.setDocRows(baseDataTable);
		baseTableContent.setJsonAble(true);
		baseTableContent.setObj(obj);
		baseTableContent.setMerges(Arrays.asList(new RowMerge(0, 3, 4), new RowMerge(0, 6, 7), new RowMerge(0, 8, 12)));

		PoiDocExport docExport = new PoiDocExport();
		docExport.export(areaContent);
		docExport.export(baseTableContent);

		exportAssessResultContent(docExport, obj, true);
		
		exportProblemDescription(docExport, obj, true);
		
		exportTypicalProblem(docExport, obj, true);
		
		exportHeader(docExport);

		docExport.write("C:/Users/xing/Desktop/" + UUID.randomUUID() + ".docx");
	}
	
	private static void exportHeader(PoiDocExport docExport) {
		List<DocHeaderRowContent> docRows = new LinkedList<>();
		DocHeaderRowContent row = new DocHeaderRowContent();
		
		String path = "C:/Users/xing/Desktop/header.png";
		
		Style style1 = new Style();
		style1.setFontType("黑体").setFontSize(10);
		style1.setHeight(48);
		style1.setWidth(184);
		row.addRowColumn(new DocColumnContent(0, AlignmentEnum.BOTTOM_OR_RIGHT).addDocText(new DocTextContent(0, true, path, style1).setPictureAble(true)));
		docRows.add(row);
		
		DocAreaContent areaContent = new DocAreaContent();
		areaContent.setDocRows(docRows);
		
		docExport.export(areaContent);
	}

	public static List<DocRowContent> test() {
		List<DocRowContent> docRows = new LinkedList<>();

		Style titleStyle = getTitleStyle();

		Style style1 = new Style();
		style1.setFontType("黑体").setFontSize(25);
		style1.setFontColor("FF0000");
		docRows.add(new DocRowContent(0)
				.addRowColumn(DocColumnContent.centerColumn(0).addDocText(new DocTextContent(0, "name", style1))));

		docRows.add(new DocRowContent(1)
				.addRowColumn(DocColumnContent.centerColumn(0).addDocText(new DocTextContent(0, "biaoduan", style1))));

		Style style2 = new Style();
		style2.setFontType("华文行楷").setFontSize(35);

		docRows.add(new DocRowContent(2).setBreakLineNumber(2)
				.addRowColumn(DocColumnContent.centerColumn(0).addDocText(new DocTextContent(0, true, "工程品质咨询报告", style2))));

		Style style3 = new Style();
		style3.setFontType("宋体").setFontSize(15).setBoldAble(true);

		docRows.add(new DocRowContent(3).setBreakLineNumber(5).addRowColumn(new DocColumnContent(0).addDocText(
				new DocTextContent(0, true, "\t\t\t报告编号：", style3), new DocTextContent(2, "code", style3))));

		docRows.add(new DocRowContent(4).addRowColumn(new DocColumnContent(0).addDocText(
				new DocTextContent(0, true, "\t\t\t建设单位：", style3), new DocTextContent(2, "test1", style3))));

		docRows.add(new DocRowContent(5).addRowColumn(new DocColumnContent(0).addDocText(
				new DocTextContent(0, true, "\t\t\t施工单位：", style3), new DocTextContent(2, "test1", style3))));

		docRows.add(new DocRowContent(6).addRowColumn(new DocColumnContent(0).addDocText(
				new DocTextContent(0, true, "\t\t\t监理单位：", style3), new DocTextContent(2, "test1", style3))));

		docRows.add(new DocRowContent(7).addRowColumn(
				new DocColumnContent(0).addDocText(new DocTextContent(0, true, "\t\t\t评估单位：xxxxxx有限公司", style3))));

		docRows.add(new DocRowContent(8).addRowColumn(new DocColumnContent(0).addDocText(
				new DocTextContent(0, true, "\t\t\t编制日期：", style3), new DocTextContent(2, "test1", style3))));

		docRows.add(new DocRowContent(9).setBreakLineNumber(2).addRowColumn(
				new DocColumnContent(0).addDocText(new DocTextContent(0, true, "第一部分： 基本资料", titleStyle))));

		return docRows;
	}

	private static Style getTableStyle(String color) {
		Style style = new Style();
		style.setFontType("宋体");
		style.setFontSize(11);
		style.setBoldAble(true);
		style.setFontColor(color);
		return style;
	}

	private static Style getTableStyle() {
		return getTableStyle(null);
	}

	/**
	 * 
	 * <p>
	 * 大标题
	 * </P>
	 * 
	 * @return
	 */
	private static Style getTitleStyle() {
		Style titleStyle = new Style();
		titleStyle.setFontType("宋体").setFontSize(20);
		return titleStyle;
	}

	/**
	 * 
	 * <p>
	 * 小标题
	 * </P>
	 * 
	 * @return
	 */
	private static Style getMinTitleStyle() {
		Style titleStyle = new Style();
		titleStyle.setFontType("宋体").setFontSize(15);
		return titleStyle;
	}

	/**
	 * 
	 * <p>
	 * 正文
	 * </P>
	 * 
	 * @return
	 */
	private static Style getTextStyle() {
		Style titleStyle = new Style();
		titleStyle.setFontType("宋体").setFontSize(11);
		return titleStyle;
	}

	private static List<DocRowContent> getBaseDataTable() {
		String font = "宋体";
		Style style = new Style();
		style.setFontType(font).setFontSize(12);
		style.setHeight(600);
		List<DocRowContent> docRows = new LinkedList<>();

		Integer clo0Width = 1500;
		Integer clo1Width = 1000;
		Integer clo2Width = 2000;
		Integer clo3Width = 2000;
		Integer clo4Width = 3000;
		
		ColumnStyle clo0 = ColumnStyle.centerStyle(clo0Width, color);
		ColumnStyle clo1 = ColumnStyle.centerStyle(clo1Width);
		ColumnStyle clo2 = ColumnStyle.centerStyle(clo2Width);
		ColumnStyle clo3 = ColumnStyle.centerStyle(clo3Width);
		ColumnStyle clo4 = ColumnStyle.centerStyle(clo4Width);
		
		Style tableStyle = getTableStyle();

		docRows.add(new DocTableRowContent(0, 600).addMerge(new ColumnMerge(1, 2)).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "项目名称", tableStyle)),
				new DocTableColumnContent(1, clo1).addDocText(new DocTextContent(0, "test1", style)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(3, clo3)
						.addDocText(new DocTextContent(0, true, "评估日期", tableStyle)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, "test1", style))));

		docRows.add(new DocTableRowContent(1, 600).addMerge(new ColumnMerge(1, 2)).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "所属区域", tableStyle)),
				new DocTableColumnContent(1, clo1).addDocText(new DocTextContent(0, "test1", style)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(3, clo3)
						.addDocText(new DocTextContent(0, true, "产品类型", tableStyle)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, "biaoduan", style))));// 标段类型

		docRows.add(new DocTableRowContent(2, 600).addMerge(new ColumnMerge(1, 2)).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "评估类型", tableStyle)),
				new DocTableColumnContent(1, clo1).addDocText(new DocTextContent(0, "biaoduan", style)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(3, ColumnStyle.centerStyle(clo3Width, color))
						.addDocText(new DocTextContent(0, true, "评估轮次", tableStyle)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, "biaoduan", style))));

		docRows.add(new DocTableRowContent(3, 600).addMerge(new ColumnMerge(1, 2)).addMerge(new ColumnMerge(3, 4))
				.addRowColumn(new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "评估人员", tableStyle)),
						new DocTableColumnContent(1, clo1).addDocText(new DocTextContent(0, true, "组长", style)),
						new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, true, "", style)),
						new DocTableColumnContent(3, clo3).addDocText(new DocTextContent(0, true, "组员", style)),
						new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, true, "", style))));

		docRows.add(new DocTableRowContent(4, 600).addMerge(new ColumnMerge(1, 2)).addMerge(new ColumnMerge(3, 4))
				.addRowColumn(new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "评估人员", tableStyle)),
						new DocTableColumnContent(1, clo1).addDocText(new DocTextContent(0, "user.name", style)),
						new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, true, "", style)),
						new DocTableColumnContent(3, clo4).addDocText(new DocTextContent(0, true, "/", style)),
						new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, true, "", style))));

		docRows.add(new DocTableRowContent(5, 600).addMerge(new ColumnMerge(1, 2)).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "监理单位", tableStyle)),
				new DocTableColumnContent(1, clo1).addDocText(new DocTextContent(0, "biaoduan", style)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(3, clo3)
						.addDocText(new DocTextContent(0, true, "标段", tableStyle)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, "biaoduan", style))));

		docRows.add(new DocTableRowContent(6, 600).addMerge(new ColumnMerge(2, 3)).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "施工单位", tableStyle)),
				new DocTableColumnContent(1, ColumnStyle.centerStyle(clo1Width, color))
						.addDocText(new DocTextContent(0, true, "承包性质", tableStyle)),
				new DocTableColumnContent(2, ColumnStyle.centerStyle(clo2Width, color))
						.addDocText(new DocTextContent(0, true, "单位名称", tableStyle)),
				new DocTableColumnContent(3, clo3).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(4, ColumnStyle.centerStyle(clo4Width, color))
						.addDocText(new DocTextContent(0, true, "承包范围", tableStyle))));

		docRows.add(new DocTableRowContent(7, 600).addMerge(new ColumnMerge(2, 3)).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "施工单位", tableStyle)),
				new DocTableColumnContent(1, clo1).addDocText(new DocTextContent(0, "chengbaoxingzhi", style)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, "danwei", style)),
				new DocTableColumnContent(3, clo3).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, "biaoduan", style))));

		docRows.add(new DocTableRowContent(8, 600).addMerge(new ColumnMerge(2, 4)).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "项目简介", tableStyle)),
				new DocTableColumnContent(1, ColumnStyle.centerStyle(clo1Width, color))
						.addDocText(new DocTextContent(0, true, "项目地址", tableStyle)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, "address", style)),
				new DocTableColumnContent(3, clo3).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, true, "", style))));

		docRows.add(new DocTableRowContent(9, 600).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "项目简介", tableStyle)),
				new DocTableColumnContent(1, ColumnStyle.centerStyle(clo1Width, color))
						.addDocText(new DocTextContent(0, true, "建筑面积（㎡）", tableStyle)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, "mianji", style)),
				new DocTableColumnContent(3, clo2).addDocText(new DocTextContent(0, true, "项目负责人", style)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, "fuzeren", style))));

		docRows.add(new DocTableRowContent(10, 600).addMerge(new ColumnMerge(2, 4)).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "项目简介", tableStyle)),
				new DocTableColumnContent(1, ColumnStyle.centerStyle(clo1Width, color))
						.addDocText(new DocTextContent(0, true, "项目组成", tableStyle)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, "biaoduan", style)),
				new DocTableColumnContent(3, clo3).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, true, "", style))));

		docRows.add(new DocTableRowContent(11, 600).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "项目简介", tableStyle)),
				new DocTableColumnContent(1, ColumnStyle.centerStyle(clo1Width, color))
						.addDocText(new DocTextContent(0, true, "结构形式", tableStyle)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, true, "/", style)),
				new DocTableColumnContent(3, clo2).addDocText(new DocTextContent(0, true, "交付日期", style)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, true, "/", style))));

		docRows.add(new DocTableRowContent(12, 600).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "项目简介", tableStyle)),
				new DocTableColumnContent(1, ColumnStyle.centerStyle(clo1Width, color))
						.addDocText(new DocTextContent(0, true, "交付标准", tableStyle)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, "jiaofuleixng", style)),
				new DocTableColumnContent(3, clo2).addDocText(new DocTextContent(0, true, "已交付情况", style)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, true, "/", style))));

		docRows.add(new DocTableRowContent(13, 600).addMerge(new ColumnMerge(1, 4)).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "项目进度", tableStyle)),
				new DocTableColumnContent(1, clo1).addDocText(new DocTextContent(0, true, "/", style)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(3, clo3).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, true, "", style))));

		docRows.add(new DocTableRowContent(14, 600).addMerge(new ColumnMerge(1, 4)).addRowColumn(
				new DocTableColumnContent(0, clo0).addDocText(new DocTextContent(0, true, "备    注", tableStyle)),
				new DocTableColumnContent(1, clo1).addDocText(new DocTextContent(0, true, "/", style)),
				new DocTableColumnContent(2, clo2).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(3, clo3).addDocText(new DocTextContent(0, true, "", style)),
				new DocTableColumnContent(4, clo4).addDocText(new DocTextContent(0, true, "", style))));

		return docRows;
	}

	/**
	 * 
	 * <p>
	 * 第二部分
	 * </P>
	 * 
	 * @param docExport
	 */
	public static void exportAssessResultContent(PoiDocExport docExport, Object obj, boolean isJson) {
		List<DocRowContent> docRows = new LinkedList<>();
		docRows.add(new DocRowContent(0).setBreakLineNumber(2).addRowColumn(
				new DocColumnContent(0).addDocText(new DocTextContent(0, true, "第二部分： 评估结果", titleStyle))));
		docRows.add(new DocRowContent(1).addRowColumn(
				new DocColumnContent(0).addDocText(new DocTextContent(0, true, "一、综合评估结果", minTitleStyle))));

		DocAreaContent titleContent = new DocAreaContent();
		titleContent.setDocRows(docRows);
		titleContent.setJsonAble(true);
		titleContent.setObj(obj);

		docExport.export(titleContent);

		// -----------------一、综合评估结果-----表格----------------------------

		List<DocRowContent> docTableRows = new LinkedList<>();

		docTableRows.add(new DocTableRowContent(0, DEFAULT_HEIGHT).addMerge(new ColumnMerge(0, 2)).addRowColumn(
				new DocTableColumnContent(0, tableCenterStyle)
						.addDocText(new DocTextContent(0, true, "评估成绩汇总", boldTextStyle)),
				new DocTableColumnContent(1, tableCenterStyle).addDocText(new DocTextContent(0, true, "", textStyle)),
				new DocTableColumnContent(2, tableCenterStyle).addDocText(new DocTextContent(0, true, "", textStyle))));

		docTableRows.add(new DocTableRowContent(1, DEFAULT_HEIGHT).addMerge(new ColumnMerge(1, 2)).addRowColumn(
				new DocTableColumnContent(0, tableCenterStyle)
						.addDocText(new DocTextContent(0, true, "项目名称", textStyle)),
				new DocTableColumnContent(1, tableCenterStyle)
						.addDocText(new DocTextContent(0, "projectName", textStyle)),
				new DocTableColumnContent(2, tableCenterStyle).addDocText(new DocTextContent(0, true, "", textStyle))));

		docTableRows.add(new DocTableRowContent(2, DEFAULT_HEIGHT).addRowColumn(
				new DocTableColumnContent(0, tableCenterStyle)
						.addDocText(new DocTextContent(0, true, "检查维度", textStyle)),
				new DocTableColumnContent(1, tableCenterStyle).addDocText(new DocTextContent(0, true, "权重", textStyle)),
				new DocTableColumnContent(2, tableCenterStyle)
						.addDocText(new DocTextContent(0, true, "得分", textStyle))));

		docTableRows.add(new DocTableRowContent(3, DEFAULT_HEIGHT).addRowColumn(
				new DocTableColumnContent(0, tableCenterStyle)
						.addDocText(new DocTextContent(0, true, "质量专项", textStyle)),
				new DocTableColumnContent(1, tableCenterStyle).addDocText(new DocTextContent(0, "aaaa", textStyle)
						.setDefaultContent("/").setSuffixContent("%").setNoValueSuffixAble(false)),
				new DocTableColumnContent(2, tableCenterStyle).addDocText(new DocTextContent(0, "aaaa", textStyle)
						.setDefaultContent("/").setSuffixContent("%").setNoValueSuffixAble(false))));

		docTableRows.add(new DocTableRowContent(4, DEFAULT_HEIGHT).addRowColumn(
				new DocTableColumnContent(0, tableCenterStyle)
						.addDocText(new DocTextContent(0, true, "安全专项", textStyle)),
				new DocTableColumnContent(1, tableCenterStyle).addDocText(new DocTextContent(0, "aaaa", textStyle)
						.setDefaultContent("/").setSuffixContent("%").setNoValueSuffixAble(false)),
				new DocTableColumnContent(2, tableCenterStyle).addDocText(new DocTextContent(0, "aaaa", textStyle)
						.setDefaultContent("/").setSuffixContent("%").setNoValueSuffixAble(false))));

		docTableRows.add(new DocTableRowContent(5, DEFAULT_HEIGHT).addMerge(new ColumnMerge(1, 2)).addRowColumn(
				new DocTableColumnContent(0, tableCenterStyle)
						.addDocText(new DocTextContent(0, true, "加扣总分", textStyle)),
				new DocTableColumnContent(1, tableCenterStyle).addDocText(new DocTextContent(0, true, "/", textStyle)),
				new DocTableColumnContent(2, tableCenterStyle).addDocText(new DocTextContent(0, true, "", textStyle))));

		docTableRows.add(new DocTableRowContent(6, DEFAULT_HEIGHT).addMerge(new ColumnMerge(1, 2)).addRowColumn(
				new DocTableColumnContent(0, tableCenterStyle)
						.addDocText(new DocTextContent(0, true, "综合得分", textStyle)),
				new DocTableColumnContent(1, tableCenterStyle).addDocText(new DocTextContent(0, "score", textStyle)),
				new DocTableColumnContent(2, tableCenterStyle).addDocText(new DocTextContent(0, true, "", textStyle))));

		docTableRows.add(new DocTableRowContent(7, DEFAULT_HEIGHT).addMerge(new ColumnMerge(1, 2)).addRowColumn(
				new DocTableColumnContent(0, tableCenterStyle).addDocText(new DocTextContent(0, true, "备注", textStyle)),
				new DocTableColumnContent(1, tableCenterStyle)
						.addDocText(new DocTextContent(0, true, "综合成绩=质量专项*30.00%+安全专项*70.00%", textStyle)),
				new DocTableColumnContent(2, tableCenterStyle).addDocText(new DocTextContent(0, true, "", textStyle))));

		DocTableAreaContent tableContent = new DocTableAreaContent();
		tableContent.setDocRows(docTableRows);
		tableContent.setJsonAble(true);
		tableContent.setObj(obj);

		docExport.export(tableContent);

		// -----------------二、分项评估结果-----表格----------------------------

		List<DocRowContent> docRows2 = new LinkedList<>();
		docRows2.add(new DocRowContent(0).setBreakLineNumber(2)
				.addRowColumn(new DocColumnContent(0).addDocText(new DocTextContent(0, true, "二、分项评估结果", minTitleStyle))));
		docRows2.add(new DocRowContent(1).addRowColumn(
				new DocColumnContent(0).addDocText(new DocTextContent(0, true, "1、质量专项", minTitleStyle))));
		docRows2.add(new DocRowContent(2).addRowColumn(
				new DocColumnContent(0).addDocText(new DocTextContent(0, true, "1.1 主要检查内容得分率", textStyle))));

		DocAreaContent titleContent2 = new DocAreaContent();
		titleContent2.setDocRows(docRows2);
		titleContent2.setJsonAble(true);
		titleContent2.setObj(obj);

		docExport.export(titleContent2);

		// 质量风险表格
		exportItemTable(docExport, obj, isJson);

		// -----------------2、安全专项合格率----------------------------

		List<DocRowContent> docRows3 = new LinkedList<>();
		docRows3.add(new DocRowContent(0).addRowColumn(
				new DocColumnContent(0).addDocText(new DocTextContent(0, true, "2、安全专项合格率", minTitleStyle))));
		docRows3.add(new DocRowContent(1).addRowColumn(
				new DocColumnContent(0).addDocText(new DocTextContent(0, true, "2.1主要检查内容得分率", textStyle))));

		DocAreaContent titleContent3 = new DocAreaContent();
		titleContent3.setDocRows(docRows3);
		titleContent3.setJsonAble(true);
		titleContent3.setObj(obj);

		docExport.export(titleContent3);

		// 安全专项表格
		exportItemTable(docExport, obj, isJson);
	}

	/**
	 * 
	 * <p>
	 * 专项表格导出
	 * </P>
	 * 
	 * @param docExport
	 * @param obj
	 * @param isJson
	 */
	public static void exportItemTable(PoiDocExport docExport, Object obj, boolean isJson) {
		List<DocRowContent> docTableRows2_1 = new LinkedList<>();
		
		Integer clo0Width = 2000;
		Integer clo1Width = 6000;
		Integer clo2Width = 2000;
		ColumnStyle style0 = ColumnStyle.centerStyle(clo0Width, color);
		ColumnStyle style1 = ColumnStyle.centerStyle(clo1Width, color);
		ColumnStyle style2 = ColumnStyle.centerStyle(clo2Width, color);
		
		docTableRows2_1.add(new DocTableRowContent(0, DEFAULT_HEIGHT).addRowColumn(
				new DocTableColumnContent(0, style0)
						.addDocText(new DocTextContent(0, true, "检查大项", boldTextStyle)),
				new DocTableColumnContent(1, style1)
						.addDocText(new DocTextContent(0, true, "检查项", boldTextStyle)),
				new DocTableColumnContent(2, style2)
						.addDocText(new DocTextContent(0, true, "分项得分", boldTextStyle))));

		DocTableAreaContent tableContent2_1 = new DocTableAreaContent();
		tableContent2_1.setDocRows(docTableRows2_1);
		tableContent2_1.setJsonAble(true);

		docExport.export(tableContent2_1);

		List<DocRowContent> docTableRows2_2 = new LinkedList<>();
		docTableRows2_2.add(new DocTableRowContent(1, DEFAULT_HEIGHT).addRowColumn(
				new DocTableColumnContent(0, style0)
						.addDocText(new DocTextContent(0, "a1", textStyle)),
				new DocTableColumnContent(1, style1)
						.addDocText(new DocTextContent(0, "a2", textStyle)),
				new DocTableColumnContent(2, style2)
						.addDocText(new DocTextContent(0, "a3", textStyle).setDefaultContent("/").setSuffixContent("%")
								.setNoValueSuffixAble(false))));

		JSONArray arr = new JSONArray();
		JSONObject json1 = new JSONObject();
		json1.put("a1", "渗漏专项");
		json1.put("a2", "屋面渗漏");
		json1.put("a3", "10.00");
		arr.add(json1);
		arr.add(json1);
		arr.add(json1);
		arr.add(json1);

		DocTableAreaContent tableContent2_2 = new DocTableAreaContent();
		tableContent2_2.setDocRows(docTableRows2_2);
		tableContent2_2.setJsonAble(true);
		tableContent2_2.setObj(arr);

		docExport.export(tableContent2_2);

		List<DocRowContent> docTableRows2_3 = new LinkedList<>();
		docTableRows2_3.add(new DocTableRowContent(1, DEFAULT_HEIGHT).addMerge(new ColumnMerge(0, 1)).addRowColumn(
				new DocTableColumnContent(0, style0)
						.addDocText(new DocTextContent(0, true, "综合得分", textStyle)),
				new DocTableColumnContent(1, style1)
						.addDocText(new DocTextContent(0, true, "综合得分", textStyle)),
				new DocTableColumnContent(2, style2)
						.addDocText(new DocTextContent(0, "score", textStyle).setDefaultContent("/")
								.setSuffixContent("%").setNoValueSuffixAble(false))));
		DocTableAreaContent tableContent2_3 = new DocTableAreaContent();
		tableContent2_3.setDocRows(docTableRows2_3);
		tableContent2_3.setJsonAble(true);
		tableContent2_3.setObj(obj);

		docExport.export(tableContent2_3);
	}
	
	/**
	 * 
	 * <p>
	 * 问题描述导出
	 * </P>
	 */
	public static void exportProblemDescription(PoiDocExport docExport, Object obj, boolean isJson) {
		List<DocRowContent> docRows = new LinkedList<>();
		docRows.add(new DocRowContent(0).setBreakLineNumber(2)
				.addRowColumn(new DocColumnContent(0).addDocText(new DocTextContent(0, true, "三、主要问题\r\n描述", minTitleStyle))));
		DocAreaContent titleContent = new DocAreaContent();
		titleContent.setDocRows(docRows);
		titleContent.setJsonAble(true);
		docExport.export(titleContent);
		
		List<DocRowContent> docTableRows = new LinkedList<>();
		DocTableRowContent tableRowContent = new DocTableRowContent(0, 2000);
		DocTableColumnContent columnContent = new DocTableColumnContent(0, new ColumnStyle().setWidth(10000).setVerticalAlignment(AlignmentEnum.TOP_OR_LEFT));
		docTableRows.add(tableRowContent);
		tableRowContent.addRowColumn(columnContent);
		
		List<DocTextContent> textList = new LinkedList<DocTextContent>();
		textList.add(new DocTextContent(0, true, "1、一、二类问题：", boldTextStyle).setNewParagraph(true));
		
		textList.add(new DocTextContent(0, true, "（1）一类问题：", boldTextStyle).setNewParagraph(true));
		textList.add(new DocTextContent(0, true, "1.xxxxx", boldTextStyle).setNewParagraph(true).setPrefixContent("\t"));
		
		textList.add(new DocTextContent(0, true, "（2）二类问题：", boldTextStyle).setNewParagraph(true));
		textList.add(new DocTextContent(0, true, "1.xxxxx", boldTextStyle).setNewParagraph(true).setPrefixContent("\t"));
		
		textList.add(new DocTextContent(0, true, "（3）风险提示问题：", boldTextStyle).setNewParagraph(true));
		textList.add(new DocTextContent(0, true, "1.xxxxx", boldTextStyle).setNewParagraph(true).setPrefixContent("\t"));
		
		textList.add(new DocTextContent(0, true, "2、各维度问题情况：", boldTextStyle).setNewParagraph(true));
		textList.add(new DocTextContent(0, true, "（1）安全文明：", boldTextStyle).setNewParagraph(true));
		textList.add(new DocTextContent(0, true, "1.xxxxx", boldTextStyle).setNewParagraph(true).setPrefixContent("\t"));
		
		columnContent.setDocTexts(textList);
		DocTableAreaContent tableContent = new DocTableAreaContent();
		tableContent.setDocRows(docTableRows);
		tableContent.setJsonAble(true);
		tableContent.setObj(obj);
		docExport.export(tableContent);
	}
	
	/**
	 * 
	 * <p>
	 * 典型问题导出
	 * </P>
	 * @param docExport
	 * @param obj
	 * @param isJson
	 */
	public static void exportTypicalProblem(PoiDocExport docExport, Object obj, boolean isJson) {
		List<DocRowContent> docRows = new LinkedList<>();
		docRows.add(new DocRowContent(0).setBreakLineNumber(2)
				.addRowColumn(new DocColumnContent(0).addDocText(new DocTextContent(0, true, "第三部分：项目典型问题展示", titleStyle))));
		DocAreaContent titleContent = new DocAreaContent();
		titleContent.setDocRows(docRows);
		titleContent.setJsonAble(true);
		docExport.export(titleContent);
		
		exportTypicalProblemTable(docExport, obj, isJson, "渗漏专项");
		exportTypicalProblemTable(docExport, obj, isJson, "渗漏专项");
	}
	
	public static void exportTypicalProblemTable(PoiDocExport docExport, Object obj, boolean isJson, String itemName) {
		List<DocRowContent> docTableHeaderRows = new LinkedList<>();

		Integer clo0Width = 4000;
		Integer clo1Width = 6000;
		
		docTableHeaderRows.add(new DocTableRowContent(0, DEFAULT_HEIGHT).addMerge(new ColumnMerge(0, 1)).addRowColumn(
				new DocTableColumnContent(0, ColumnStyle.centerStyle(clo0Width, color))
						.addDocText(new DocTextContent(0, true, itemName, boldTextStyle)),
				new DocTableColumnContent(1, ColumnStyle.centerStyle(clo1Width, color))
						.addDocText(new DocTextContent(0, true, "", boldTextStyle))));
		
		DocTableAreaContent tableHeaderContent = new DocTableAreaContent();
		tableHeaderContent.setDocRows(docTableHeaderRows);
		tableHeaderContent.setJsonAble(true);
		docExport.export(tableHeaderContent);
		
		List<DocRowContent> docTableRows = new LinkedList<>();

		docTableRows.add(new DocTableRowContent(0, DEFAULT_HEIGHT).addRowColumn(
				new DocTableColumnContent(0, ColumnStyle.centerStyle(clo0Width))
						.addDocText(new DocTextContent(0, "picture", textStyle).setPictureAble(true)),
				new DocTableColumnContent(1, ColumnStyle.centerStyle(clo1Width))
						.addDocText(new DocTextContent(0, true, "11111", textStyle))));
		
		JSONArray arr = new JSONArray();
		arr.add("C:/Users/xing/Desktop/test/1.png");
		arr.add("C:/Users/xing/Desktop/test/1.png");
		arr.add("C:/Users/xing/Desktop/test/1.png");
		arr.add("C:/Users/xing/Desktop/test/1.png");
		arr.add("C:/Users/xing/Desktop/test/1.png");
		JSONObject json = new JSONObject();
		json.put("picture", arr);
		
		DocTableAreaContent tableContent = new DocTableAreaContent();
		tableContent.setDocRows(docTableRows);
		tableContent.setJsonAble(true);
		tableContent.setObj(json);
		docExport.export(tableContent);
	}
}
