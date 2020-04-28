package cn._100wb.excel.service;

import java.util.List;
import java.util.Map;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 服务层通用工具接口
 * @author WanTD
 * @time 2017年3月23日上午9:50:13
 */
public interface ServiceUtil {
	/**
	 * 给一个excel,创建一个sheet
	 * 2017年3月23日上午10:42:47
	 * @param listdata 数据集合
	 * @param workbook excle对象
	 * @param titles   表头名称
	 * @param colums   指定列值,对应bean对象的属性名称
	 * @param excleName sheet的名称
	 * @return 
	 * XSSFWorkbook   工作表
	 */
	XSSFWorkbook createSheet(List<?> listdata, XSSFWorkbook workbook, String[] titles, String[] colums, String excleName);

	/**
	 * 构建EXCEL,创建多个SHEET
	 * @param list_list_mapDate 数据集合
	 * @param workbook excle对象
	 * @param titles   表头名称
	 * @param colums   指定列值,对应bean对象的属性名称
	 * @param excleNameArray sheet的名称数组
	 * @author WanTD
	 * @version 2017年5月15日 下午4:35:22
	 * @return XSSFWorkbook
	 */
	XSSFWorkbook createSheet(List<List<Map<String, Object>>> list_list_mapDate, XSSFWorkbook workbook, String[] titles, String[] colums, String[] excleNameArray) throws Exception;
	
	/**
	  * 设置Excel表头列宽
	  * @author WanTD
	  * @version 2018年10月25日 下午2:49:23
	  * @return void
	  */
	void setColumnStyle(XSSFSheet sheet, int k, int columnWidth) throws Exception;
	
	/**
	  * 为多个Sheet创建单元格,扩展重构单元格流程
	  * @author WanTD
	  * @version 2018年12月25日 上午9:13:47
	  * @return void
	  */
	void creatSheetArrayCell(XSSFCellStyle shortStyle, String value, XSSFRow row, int j) throws Exception;
	
	/**
	  * 为单个个Sheet创建单元格,扩展重构单元格流程
	  * @author WanTD
	  * @version 2018年12月25日 上午9:12:45
	  * @return void
	  */
	void creatSheetCell(XSSFCellStyle shortStyle, String value, XSSFRow row, int j) throws Exception;
	
}
