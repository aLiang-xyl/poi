package com.framework.common;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;
import java.util.UUID;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.framework.poi.export.excel.PictureInfo;
import com.framework.poi.export.excel.PoiExport;
import com.framework.poi.export.excel.PoiExportSheet;
import com.framework.poi.export.excel.entity.ExcelColumn;

import lombok.Data;

/**
 * <p>
 * 描述: 人员导出测试
 * </p>
 * 
 * @author aLiang
 * @date 2020年8月20日 下午4:33:37
 */
@Data
public class Person {

	private Long id;

	private String name;

	private LocalDate birthday;

	private Integer sex;

	private String address;

	private LocalDateTime createTime;

	private Company company;

//	@ExcelExport(order = 6, isPicture = true, width = 40, name = "头像")
//	private String headPic;
//	
	private String pic;
	
	private PictureInfo pic1;
	
	private List<PictureInfo> listPic;
	
	private List<String> workPic;

	public static void main(String[] args) {
		List<Person> list = new ArrayList<>();
		for (int i = 0; i < 10; i++) {
			Person p1 = new Person();

			Company company = new Company();
			if (i < 5) {
				p1.setId(1L);
				company.setName("公司" + 1);
				company.setId(1L);
			} else {
				p1.setId(2L);
				company.setName("公司" + 2);
				company.setId(2L);
			}

			p1.setName("李四" + i);
			p1.setBirthday(LocalDate.now());
			p1.setCreateTime(LocalDateTime.now());
			p1.setSex(i % 2);
			p1.setAddress("广东省深圳市福田区上梅林街道");
			p1.setCompany(company);
//			p1.setHeadPic("C:/Users/xing/Desktop/test/" + i + ".jpg");
			
			List<String> picList = new LinkedList<String>();
			for (int j = 0; j < 5; j++) {
				picList.add("C:/Users/xing/Desktop/test/" + j + ".png");
			}
			for (int j = 0; j < 5; j++) {
				picList.add("C:/Users/xing/Desktop/test/" + j + ".png");
			}
			
			PictureInfo pic = new PictureInfo();
			pic.setPath("C:/Users/xing/Desktop/test/" + 1 + ".png");
			
			p1.setPic(pic.getPath());
			p1.setPic1(pic);
			p1.setListPic(Arrays.asList(pic));
			p1.setWorkPic(picList);
			list.add(p1);
		}
		// xlsx格式导出
//		long l1 = System.currentTimeMillis();
//		PoiExport info = PoiExport.getInstance(list, ExcelTypeEnum.XLSX, 50, "test");
//		ExcelExportUtils.export(info, "C:/Users/xing/Desktop/" + UUID.randomUUID() + ".xlsx");
//		System.out.println(System.currentTimeMillis() - l1);

		// xls格式导出
//		long l2 = System.currentTimeMillis();
//		PoiExport info1 = PoiExport.getInstance(list, ExcelTypeEnum.XLS, 200, "test");
//		ExcelExportUtils.export(info1, "C:/Users/xing/Desktop/" + UUID.randomUUID() + ".xls");
//		System.out.println(System.currentTimeMillis() - l2);

		List<ExcelColumn> template = getTemplate();
		PoiExport p = new PoiExport();

		PoiExportSheet poiSheet = new PoiExportSheet();
		poiSheet.setColumns(template);
		poiSheet.setDescription("人员信息");
		poiSheet.setRows(list);
		poiSheet.setSheetName("人员信息");
		poiSheet.setTitleName("人员导出信息");
		poiSheet.setHeight(100);
		poiSheet.setPictureScale(new BigDecimal("0.5"));

		p.export(poiSheet);

		PoiExportSheet poiSheet1 = new PoiExportSheet();
		poiSheet1.setColumns(template);
		poiSheet1.setDescription("人员信息1");
		poiSheet1.setRows(JSONArray.parseArray(JSONObject.toJSONString(list)));
		poiSheet1.setRowsIsJson(true);
		poiSheet1.setSheetName("人员信息1");
		poiSheet1.setTitleName("人员导出信息1");
		poiSheet1.setHeight(100);
		p.export(poiSheet1);
		p.write("C:/Users/xing/Desktop/" + UUID.randomUUID() + ".xlsx");
	}

	private static List<ExcelColumn> getTemplate() {
		List<ExcelColumn> list = new LinkedList<ExcelColumn>();
		ExcelColumn t0 = new ExcelColumn();
		t0.setColumnIndex(0);
		t0.setTitleName("id");
		t0.setFieldName("id");
		t0.setMergeAble(true);
		t0.setMergeKeyField("id");
		list.add(t0);

		ExcelColumn t1 = new ExcelColumn();
		t1.setColumnIndex(1);
		t1.setTitleName("姓名");
		t1.setFieldName("name");
		list.add(t1);

		list.add(new ExcelColumn(2, "生日", "birthday").setDateTimeFromat("yyyy年MM月dd日"));
		list.add(new ExcelColumn(4, "地址", "address"));

		ExcelColumn excelExportTemplate = new ExcelColumn(5, "所属公司", "company.name");
		excelExportTemplate.setMergeAble(true);
		excelExportTemplate.setMergeKeyField("company.id");
		list.add(excelExportTemplate);

		ExcelColumn sexTemplate = new ExcelColumn(3, "性别", "sex");
		sexTemplate.setValueExchange("{\"1\":\"男\", \"0\":\"女\"}");
		list.add(sexTemplate);
		
		ExcelColumn picTemplate6 = new ExcelColumn(6, "图片1", "pic1");
		picTemplate6.setPictureAble(true);
		picTemplate6.setPictureMonopoly(true);
		picTemplate6.setPicHeight(new BigDecimal("0.25"));
		picTemplate6.setPicWidth(new BigDecimal("0.2"));
		picTemplate6.setWidth(20);
		list.add(picTemplate6);
		
		ExcelColumn picTemplate7 = new ExcelColumn(7, "图片2", "listPic");
		picTemplate7.setPictureAble(true);
		picTemplate7.setPictureMonopoly(true);
		picTemplate7.setPicHeight(new BigDecimal("0.25"));
		picTemplate7.setPicWidth(new BigDecimal("0.2"));
		picTemplate7.setWidth(20);
		list.add(picTemplate7);
		
		ExcelColumn picTemplate8 = new ExcelColumn(8, "图片0000", "pic");
		picTemplate8.setPictureAble(true);
		picTemplate8.setPictureMonopoly(true);
		picTemplate8.setPicHeight(new BigDecimal("0.25"));
		picTemplate8.setPicWidth(new BigDecimal("0.2"));
		picTemplate8.setWidth(20);
		list.add(picTemplate8);

		ExcelColumn picTemplate = new ExcelColumn(9, "图片3", "workPic");
		picTemplate.setPictureAble(true);
		picTemplate.setPictureMonopoly(true);
		picTemplate.setPicHeight(new BigDecimal("0.25"));
		picTemplate.setPicWidth(new BigDecimal("0.2"));
		picTemplate.setWidth(20);
		list.add(picTemplate);

		return list;
	}
}
