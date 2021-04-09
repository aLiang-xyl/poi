package com.framework.poi.export.excel;

import java.beans.IntrospectionException;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import com.alibaba.fastjson.JSONObject;
import com.framework.poi.export.Assert;
import com.framework.poi.export.BusinessException;
import com.framework.poi.export.PoiUtils;
import com.framework.poi.export.excel.entity.ExcelColumn;

import lombok.extern.log4j.Log4j2;
import net.coobird.thumbnailator.Thumbnails;
import net.coobird.thumbnailator.Thumbnails.Builder;

/**
 * <p>
 * 描述:excel导出
 * </p>
 * 
 * @author aLiang
 * @date 2020年8月20日 下午4:15:38
 */
@Log4j2
public class PoiExport {

	public static final String XLSX = ".xlsx";
	private static final String DEFAULT_SHEET_NAME = "sheet";
	private static final int DEFAULT_HEIGHT_WIDTH = -1;

	private static final Set<String> IMAGE_TYPE;

	static {
		String imageType = "bmp,jpg,png,tif,gif,pcx,tga,exif,fpx,svg,psd,cdr,pcd,dxf,ufo,eps,ai,raw,wmf,webp";
		IMAGE_TYPE = new HashSet<String>();
		for (String type : imageType.split(",")) {
			IMAGE_TYPE.add(type);
		}
	}

	private Workbook workbook;
	private Map<String, Integer> sheetRowNumberMap;
	private Map<String, Sheet> sheetMap;

	/**
	 * 
	 * <p>
	 * 构造函数
	 * </P>
	 * 
	 */
	public PoiExport() {
		super();
		this.sheetRowNumberMap = new HashMap<>();
		this.sheetMap = new HashMap<>();
		this.workbook = new XSSFWorkbook();
	}

	/**
	 * 
	 * <p>
	 * 导出到指定的sheet
	 * </P>
	 * 
	 * @param poiSheet
	 * @return
	 */
	public PoiExport export(PoiExportSheet poiSheet) {
		poiSheet.validate();
		String sheetName = poiSheet.getSheetName();
		if (StringUtils.isEmpty(sheetName)) {
			sheetName = DEFAULT_SHEET_NAME;
			poiSheet.setSheetName(DEFAULT_SHEET_NAME);
		}

		Integer separateNumber = poiSheet.getSeparateNumber();
		if (separateNumber != null) {
			int rownum = getSheetRowNumber(sheetName);
			rownum += separateNumber;
			setSheetRowNumber(sheetName, rownum);
		}

		Boolean onlyAppend = poiSheet.getOnlyAppend();
		if (onlyAppend == null || !onlyAppend) {
			createTitleRow(poiSheet);
		}

		Collection<?> list = poiSheet.getRows();
		if (list == null) {
			return this;
		}
		Iterator<?> iterator = list.iterator();
		int rownum = getSheetRowNumber(sheetName);
		while (iterator.hasNext()) {
			Object object = iterator.next();
			if (object == null) {
				continue;
			}
			try {
				createRow(rownum, object, poiSheet);
				rownum++;
			} catch (Exception e) {
				log.error("导出一行出错", e);
			}
		}
		this.setSheetRowNumber(sheetName, rownum);
		// 合并
		merge(poiSheet);
		return this;
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
			this.workbook.write(outputStream);
		} catch (IOException e) {
			log.error("导出一行出错", e);
		}
	}

	/**
	 * 
	 * <p>
	 * 获取sheet行数
	 * </P>
	 * 
	 * @param sheetName
	 * @return
	 */
	private Integer getSheetRowNumber(String sheetName) {
		sheetName = getSheetName(sheetName);
		Integer sheetRowNumber = sheetRowNumberMap.get(sheetName);
		if (sheetRowNumber == null) {
			sheetRowNumber = 0;
			setSheetRowNumber(sheetName, sheetRowNumber);
		}
		return sheetRowNumber;
	}

	/**
	 * 
	 * <p>
	 * 记录sheet行数
	 * </P>
	 * 
	 * @param sheetName
	 * @param sheetRowNumber
	 */
	private void setSheetRowNumber(String sheetName, Integer sheetRowNumber) {
		sheetName = getSheetName(sheetName);
		sheetRowNumberMap.put(sheetName, sheetRowNumber);
	}

	/**
	 * 
	 * <p>
	 * 根据sheet名称获取sheet
	 * </P>
	 * 
	 * @param sheetName
	 * @return
	 */
	public Sheet getSheet(String sheetName) {
		sheetName = getSheetName(sheetName);
		Sheet sheet = sheetMap.get(sheetName);
		if (sheet == null) {
			sheet = workbook.createSheet(sheetName);
			sheetMap.put(sheetName, sheet);
		}
		return sheet;
	}

	/**
	 * 
	 * <p>
	 * 如果没有sheetname则使用默认
	 * </P>
	 * 
	 * @param sheetName
	 * @return
	 */
	private String getSheetName(String sheetName) {
		if (StringUtils.isEmpty(sheetName)) {
			return DEFAULT_SHEET_NAME;
		}
		return sheetName;
	}

	/**
	 * 
	 * <p>
	 * 标题行
	 * </P>
	 * 
	 * @param poiExport
	 * @param poiSheet
	 * @return
	 */
	private PoiExport createTitleRow(PoiExportSheet poiSheet) {
		List<ExcelColumn> list = poiSheet.getColumns();
		String sheetName = poiSheet.getSheetName();
		Integer sheetRowNumber = this.getSheetRowNumber(sheetName);

		Sheet sheet = this.getSheet(sheetName);
		// 标题行
		String titleName = poiSheet.getTitleName();
		if (!StringUtils.isEmpty(titleName)) {
			Row titleRow = sheet.createRow(sheetRowNumber);
			titleRow.createCell(0).setCellValue(titleName);

			CellRangeAddress region = new CellRangeAddress(sheetRowNumber, sheetRowNumber, 0, list.size() - 1);
			sheet.addMergedRegion(region);

			sheetRowNumber++;
		}

		// 描述行
		String description = poiSheet.getDescription();
		if (!StringUtils.isEmpty(description)) {
			Row descriptionRow = sheet.createRow(sheetRowNumber);
			descriptionRow.createCell(0).setCellValue(description);

			CellRangeAddress region = new CellRangeAddress(sheetRowNumber, sheetRowNumber, 0, list.size() - 1);
			sheet.addMergedRegion(region);

			sheetRowNumber++;
		}

		Row row = sheet.createRow(sheetRowNumber);
		for (int i = 0; i < list.size(); i++) {
			ExcelColumn col = list.get(i);
			row.createCell(col.getColumnIndex()).setCellValue(col.getTitleName());
		}
		sheetRowNumber++;

		this.setSheetRowNumber(sheetName, sheetRowNumber);

		return this;
	}

	/**
	 * 
	 * <p>
	 * 导出一行
	 * </P>
	 * 
	 * @param rownum
	 * @param rowObj
	 * @param poiSheet
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 * @throws IntrospectionException
	 * @throws InvocationTargetException
	 */
	private void createRow(int rownum, Object rowObj, PoiExportSheet poiSheet) {
		String sheetName = poiSheet.getSheetName();
		Sheet sheet = this.getSheet(sheetName);
		Row row = sheet.createRow(rownum);
		if (rowObj instanceof PoiExportRow) {
			PoiExportRow exportRow = (PoiExportRow) rowObj;
			row.setHeightInPoints(exportRow.getHeight());
		} else {
			row.setHeightInPoints(poiSheet.getHeight());
		}

		List<ExcelColumn> columns = poiSheet.getColumns();
		for (ExcelColumn column : columns) {
			Integer columnIndex = column.getColumnIndex();
			String fieldName = column.getFieldName();
			Object value = getValue(rowObj, fieldName, poiSheet.isRowsIsJson());

			handleMerge(poiSheet, column, rowObj, rownum);

			Integer width = column.getWidth();
			width = width == null ? DEFAULT_HEIGHT_WIDTH : width;
			if (width != DEFAULT_HEIGHT_WIDTH) {
				sheet.setColumnWidth(columnIndex, width * 256);
			}
			column.setWidth(width);

			// 处理图片导出，图片路径可以是字符串类型（路径信息）， 也可以是PictureInfo类型包含描述信息
			Boolean pictureAble = column.getPictureAble();
			pictureAble = pictureAble == null ? false : pictureAble;
			if (pictureAble && value != null) {
				if (value instanceof PictureInfo || value instanceof JSONObject) {
					PictureInfo pi = null;
					if (value instanceof JSONObject) {
						JSONObject j = (JSONObject) value;
						pi = j.toJavaObject(PictureInfo.class);
					} else {
						pi = (PictureInfo) value;
					}
					if (pi == null) {
						continue;
					}
					byte[] readBytes = readBytes(pi.getPath(), poiSheet);
					pictureExport(poiSheet, pi.getDescription(), 0, rownum, readBytes, column);
				} else if (value instanceof String) {
					byte[] readBytes = readBytes(value.toString(), poiSheet);
					pictureExport(poiSheet, null, 0, rownum, readBytes, column);
				} else {
					List<?> piList = (List<?>) value;
					int count = 0;
					for (Object obj : piList) {
						if (obj == null) {
							continue;
						}
						String description = null;
						String path = null;
						if (obj instanceof String) {
							path = obj.toString();
						} else {
							PictureInfo pi = null;
							if (poiSheet.isRowsIsJson()) {
								JSONObject j = (JSONObject) obj;
								pi = j.toJavaObject(PictureInfo.class);
							} else {
								pi = (PictureInfo) obj;
							}
							if (pi == null) {
								continue;
							}
							path = pi.getPath();
							description = pi.getDescription();
						}
						byte[] readBytes = readBytes(path, poiSheet);
						if (readBytes == null) {
							continue;
						}
						pictureExport(poiSheet, description, count, rownum, readBytes, column);
						count++;
					}
				}
			}

			// 处理非图片数据导出
			else if (!pictureAble) {
				if (value == null) {
					value = column.getDefaultValue();
				}
				value = formatValue(value, column, poiSheet.isRowsIsJson());
				row.createCell(columnIndex).setCellValue(String.valueOf(value));
			}
		}
	}

	/**
	 * 
	 * <p>
	 * 合并
	 * </P>
	 * 
	 * @param poiSheet
	 */
	private void merge(PoiExportSheet poiSheet) {
		Map<Integer, PoiExportMerge> mergeMap = poiSheet.getMergeMap();
		if (CollectionUtils.isEmpty(mergeMap)) {
			return;
		}
		Sheet sheet = this.getSheet(poiSheet.getSheetName());
		Collection<PoiExportMerge> values = mergeMap.values();
		for (PoiExportMerge poiExportMerge : values) {
			Integer columnIndex = poiExportMerge.getColumnIndex();
			Collection<PoiExportMergeRow> allMerge = poiExportMerge.getAllMerge();
			for (PoiExportMergeRow merge : allMerge) {
				if (merge.getStartRowNumber() == merge.getEndRowNumber()) {
					// 如果只有一行不需要合并
					continue;
				}
				CellRangeAddress region = new CellRangeAddress(merge.getStartRowNumber(), merge.getEndRowNumber(),
						columnIndex, columnIndex);
				sheet.addMergedRegion(region);
			}
		}
	}

	/**
	 * 
	 * <p>
	 * 处理合并
	 * </P>
	 * 
	 * @param poiSheet
	 * @param column
	 * @param rowObj
	 * @param rowNumber
	 */
	private void handleMerge(PoiExportSheet poiSheet, ExcelColumn column, Object rowObj, int rowNumber) {
		Boolean mergeAble = column.getMergeAble();
		if (mergeAble == null || !mergeAble) {
			return;
		}
		String mergeKeyField = column.getMergeKeyField();
		Object mergeValue = getValue(rowObj, mergeKeyField, poiSheet.isRowsIsJson());

		Integer columnIndex = column.getColumnIndex();
		boolean containsMerge = poiSheet.containsMerge(columnIndex, mergeValue);
		if (containsMerge) {
			poiSheet.addMergeEndRow(columnIndex, mergeValue);
		} else {
			poiSheet.putMergeEndRow(columnIndex, mergeValue, rowNumber);
		}
	}

	/**
	 * 
	 * <p>
	 * 取值
	 * </P>
	 * 
	 * @param rowObj
	 * @param fieldName
	 * @return
	 */
	private Object getValue(Object rowObj, String fieldName, boolean isJson) {
		return PoiUtils.getValue(rowObj, fieldName, isJson);
	}

	/**
	 * 
	 * <p>
	 * 读取文件
	 * </P>
	 * 
	 * @param path
	 * @param poiSheet
	 * @return
	 */
	private static byte[] readBytes(String path, PoiExportSheet poiSheet) {
		if (StringUtils.isEmpty(path)) {
			return null;
		}
		ByteArrayOutputStream out = null;
		try {
			// 校验图片格式
			int lastIndexOf = path.lastIndexOf(".");
			if (lastIndexOf < 0 || lastIndexOf + 1 == path.length()
					|| !IMAGE_TYPE.contains(path.substring(lastIndexOf + 1).toLowerCase())) {
				return null;
			}
			Path pathTemp = Paths.get(path);
			if (Files.notExists(pathTemp)) {
				return null;
			}
			BigDecimal pictureScale = poiSheet.getPictureScale();
			if (pictureScale == null || pictureScale.doubleValue() <= 0) {
				pictureScale = new BigDecimal("0.5");
			}
			if (pictureScale.doubleValue() == 1) {
				// 如果不压缩
				return Files.readAllBytes(pathTemp);
			} else {
				Builder<File> scale = Thumbnails.of(pathTemp.toFile()).scale(pictureScale.doubleValue());
				out = new ByteArrayOutputStream();
				scale.toOutputStream(out);

				return out.toByteArray();
			}
		} catch (Exception e) {
			log.error("导出文件失败", e);
		} finally {
			if (out != null) {
				try {
					out.close();
				} catch (IOException e) {
					log.error("导出文件关闭流失败", e);
				}
			}
		}
		return null;
	}

	/**
	 * 
	 * <p>
	 * 导出图片处理
	 * </P>
	 * 
	 * @param poiExport
	 * @param poiSheet
	 * @param picObj
	 * @param add
	 * @param rownum
	 * @param readBytes
	 * @param column
	 */
	private void pictureExport(PoiExportSheet poiSheet, String description, int add, int rownum, byte[] readBytes,
			ExcelColumn column) {
		if (readBytes == null) {
			return;
		}
		try {
			int col = column.getColumnIndex();
			// 图片是否独占一格
			Boolean pictureMonopoly = column.getPictureMonopoly();
			if (pictureMonopoly == null) {
				pictureMonopoly = false;
			}
			Sheet sheet = this.getSheet(poiSheet.getSheetName());
			if (pictureMonopoly) {
				// 图片独占一格
				col = column.getColumnIndex() + add;
				sheet.setColumnWidth(col, column.getWidth() * 256);
			}

			boolean descriptionIsNotNull = !StringUtils.isEmpty(description);
			if (descriptionIsNotNull) {
				Row row = sheet.getRow(rownum);

				Cell cellPic = row.createCell(col);
				cellPic.setCellValue(description);

				CellStyle cellStyle = workbook.createCellStyle();
				cellStyle.setAlignment(HorizontalAlignment.CENTER);
				cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
				cellPic.setCellStyle(cellStyle);
			}

			int pictureIdx = workbook.addPicture(readBytes, Workbook.PICTURE_TYPE_JPEG);
			CreationHelper helper = workbook.getCreationHelper();
			Drawing<?> drawing = sheet.createDrawingPatriarch();
			ClientAnchor anchor = helper.createClientAnchor();

			// 图片插入坐标
			anchor.setCol1(col); // 列
			anchor.setRow1(rownum); // 行

			BigDecimal picWidth = column.getPicWidth();
			if (picWidth == null) {
				picWidth = new BigDecimal("0.999");
			}
			BigDecimal picHeight = column.getPicHeight();
			if (picHeight == null) {
				picHeight = new BigDecimal("0.999");
			}
			if (!pictureMonopoly) {
				// 单元格宽度
				int width = (int) (sheet.getColumnWidthInPixels(col));
				// 图片宽度
				int pictureWidth = (int) (width * picWidth.doubleValue());
				// 图片高度
				int pictureHeight = (int) (picHeight.doubleValue() * sheet.getRow(rownum).getHeightInPoints());
				// 每一行图片的数量
				int picRowCount = width / pictureWidth;

				// 当前图片所在行中的第n个
				int surplus = add % picRowCount;

				// 图片的行数，从第0行开始
				int picRowNumber = add / picRowCount;

				anchor.setCol2(col); // 列
				anchor.setRow2(rownum); // 行
				anchor.setDx1(Units.EMU_PER_PIXEL * surplus * pictureWidth);
				anchor.setDy1(Units.EMU_PER_PIXEL * picRowNumber * pictureHeight);
				anchor.setDx2(Units.EMU_PER_PIXEL * (surplus + 1) * pictureWidth);
				anchor.setDy2(Units.EMU_PER_PIXEL * (picRowNumber + 1) * pictureHeight);
			}

			// 插入图片
			Picture pict = drawing.createPicture(anchor, pictureIdx);
			if (descriptionIsNotNull && pictureMonopoly) {
				picHeight = picHeight.multiply(new BigDecimal("0.9"));
			}
			if (pictureMonopoly) {
				pict.resize(picWidth.doubleValue(), picHeight.doubleValue());
			} else {
				pict.resize(picWidth.doubleValue() * 0.5, picHeight.doubleValue() * 0.5);
			}

		} catch (Exception e) {
			log.error("导出文件失败", e);
		}

	}

	/**
	 * 
	 * <p>
	 * 格式化值
	 * </P>
	 * @param value
	 * @param col
	 * @param isJson
	 * @return
	 */
	private static Object formatValue(Object value, ExcelColumn col, boolean isJson) {
		return PoiUtils.formatValue(value, col.getDateTimeFromat(), col.getValueExchange(), isJson);
	}
}
