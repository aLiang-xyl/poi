package com.framework.poi.export.excel;

/**
 * <p>
 * 导出记录状态枚举
 * </P>
 * 
 * @author aLiang
 * @date 2021年1月13日 下午3:12:28
 */
public enum ExportRecordStatusEnum {

	/**
	 * 文件生成中
	 */
	IN_PROCESS(32L, "文件生成中"),
	/**
	 * 导出成功
	 */
	SUCCESS(33L, "导出成功"),
	/**
	 * 导出失败
	 */
	FAIL(34L, "导出失败"),
	/**
	 * 文件已过期
	 */
	EXPIRED(35L, "文件已过期");

	private Long status;
	private String description;

	private ExportRecordStatusEnum(Long status, String description) {
		this.status = status;
		this.description = description;
	}

	public Long getStatus() {
		return status;
	}

	public String getDescription() {
		return description;
	}

	public static ExportRecordStatusEnum valueOfStatus(Long status) {
		for (ExportRecordStatusEnum statusEnum : ExportRecordStatusEnum.values()) {
			if (statusEnum.status.equals(status)) {
				return statusEnum;
			}
		}
		return null;
	}
}
