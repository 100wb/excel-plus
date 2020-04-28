package cn._100wb.excel.service.impl;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import cn._100wb.excel.service.ServiceUtil;
import cn._100wb.excel.utils.BeanMapConvertUtil;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



import net.sf.json.JSONArray;
import net.sf.json.JSONObject;
import net.sf.json.JsonConfig;
import net.sf.json.util.CycleDetectionStrategy;

/**
 * 服务层通用工具接口实现类
 * @author WanTD
 * @time 2017年3月23日上午10:45:50
 */
public class ServiceUtilImpl implements ServiceUtil {
	
	/* (non-Javadoc)
	 * @see com.adcc.oas.service.base.ServiceUtil#createSheet(java.util.List, org.apache.poi.xssf.usermodel.XSSFWorkbook, java.lang.String[], java.lang.String[], java.lang.String)
	 */
	@Override
	public XSSFWorkbook createSheet(List<?> listdata, XSSFWorkbook workbook, String[] titles, String[] colums,
			String excleName) {
		// 数据格式化
		List<String[]> formatelist = formDatas(listdata,colums);
		// 创建表头
		XSSFSheet sheet = createTitles(excleName, workbook, titles);
		// 填入数据
		addDataToSheet(workbook, formatelist, sheet);
		return workbook;
	}
	
	/* (non-Javadoc)
	 * @see com.adcc.oas.service.sbase.ServiceUtil#createSheet(java.util.List, org.apache.poi.xssf.usermodel.XSSFWorkbook, java.lang.String[], java.lang.String[], java.lang.String[])
	 * WTD Insert
	 */
	@Override
	public XSSFWorkbook createSheet(List<List<Map<String, Object>>> list_list_mapDate, XSSFWorkbook workbook, String[] titles, String[] colums,
			String[] excleNameArray) throws Exception {
		// 解析多个SHEET数据
		List<List<String[]>> formatelist = formDataColums(list_list_mapDate,colums);
		// 创建多个SHEET标签表头
		XSSFSheet[] sheet = createTitles(excleNameArray, workbook, titles);
		// 将多个SHEET标签填充数据
		addDataToSheetArray(workbook, formatelist, sheet);
		return workbook;
	}
	

	/* (non-Javadoc)
	  * @see com.adcc.oas.service.base.ServiceUtil#setColumnStyle(org.apache.poi.xssf.usermodel.XSSFSheet, int, int)
	  * WTD Insert
	  */
	@Override
	public void setColumnStyle(XSSFSheet sheet, int k, int columnWidth){
		sheet.setColumnWidth(k, columnWidth);					
	}
	
	/* (non-Javadoc)
	  * @see com.adcc.oas.service.base.ServiceUtil#creatSheetArrayCell(org.apache.poi.xssf.usermodel.XSSFCellStyle, java.lang.String, org.apache.poi.xssf.usermodel.XSSFRow, int)
	  * WTD Insert
	  */
	public void creatSheetArrayCell(XSSFCellStyle shortStyle, String value, XSSFRow row, int j) {
		XSSFCell cell = row.createCell(j, 0);
		cell.setCellType(XSSFCell.CELL_TYPE_STRING);
		cell.setCellStyle(shortStyle);
		cell.setCellValue(value);
	}
	
	/* (non-Javadoc)
	  * @see com.adcc.oas.service.base.ServiceUtil#creatSheetCell(org.apache.poi.xssf.usermodel.XSSFCellStyle, java.lang.String, org.apache.poi.xssf.usermodel.XSSFRow, int)
	  * WTD Insert
	  */
	public void creatSheetCell(XSSFCellStyle shortStyle, String value, XSSFRow row, int j) {
		XSSFCell cell = row.createCell(j, 0);
		cell.setCellType(XSSFCell.CELL_TYPE_STRING);
		cell.setCellStyle(shortStyle);
		cell.setCellValue(value);
	}
	
	/**
	 * 向sheet表中插入数据
	 * 2017年3月23日上午11:30:26
	 * @param workbook
	 * @param formatelist
	 * @param sheet 
	 * void
	 */
	private void addDataToSheet(XSSFWorkbook workbook, List<String[]> formatelist, XSSFSheet sheet) {
		XSSFCellStyle shortStyle = workbook.createCellStyle();
		shortStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);  
		shortStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		XSSFCellStyle longStyle = workbook.createCellStyle();
		longStyle.setWrapText(true);//自动换行 
		longStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		longStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);
		for (String[] strarry:formatelist) {
			XSSFRow row = sheet.createRow(sheet.getLastRowNum()+1);
			for (int i=0;i<strarry.length;i++) {
				creatSheetCell(shortStyle, strarry[i], row, i);
			}
		}
	}

	/**
	 * 对集合数据 进行重组,让之与colums序列对应
	 * 2017年3月23日上午11:14:16
	 * @param listdata
	 * @param titles
	 * @param colums
	 * @return 
	 * List<String[]>
	 */
	@SuppressWarnings("unchecked")
	private List<String[]> formDatas(List<?> listdata,String[] colums) {
		JsonConfig jfg=new JsonConfig();
		jfg.setCycleDetectionStrategy(CycleDetectionStrategy.LENIENT);
		JSONArray jsonArray = JSONArray.fromObject(listdata, jfg);
		List<JSONObject> mapListJson = jsonArray;
		List<String[]> list=new ArrayList<>();
		for (JSONObject jobj : mapListJson) {
			String[] columValue=new String[colums.length];
			for (int i=0;i<colums.length;i++) {
				Object object = jobj.get(colums[i]);
				columValue[i]=object==null?"":object.toString();
			}
			list.add(columValue);
		}
		return list;
	}
	
	/**
	 * 创建表头
	 * 2017年3月23日上午10:46:36
	 * @param systemCode
	 * @param wb
	 * @param titles
	 * @return 
	 * XSSFSheet
	 */
	private XSSFSheet createTitles(String excleName, XSSFWorkbook wb, String[] colums) {
		XSSFFont font = wb.createFont();
		font.setColor(XSSFFont.COLOR_NORMAL);
		font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		XSSFCellStyle titleStyle = wb.createCellStyle();
		titleStyle.setFont(font);
		titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		titleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		XSSFSheet sheet = wb.createSheet(excleName);
		XSSFRow titleRow = sheet.createRow(0);
		for (int k = 0; k < colums.length; k++) {
			// 设置Excel表头列宽
			setColumnStyle(sheet, k, 8000);
			XSSFCell cell = titleRow.createCell(k, 0);
			cell.setCellStyle(titleStyle);
			cell.setCellType(XSSFCell.CELL_TYPE_STRING);
			cell.setCellValue(colums[k]);
		}
		return sheet;
	}
	
	/**
	 * 创建多标签表头
	 * @author WanTD
	 * @version 2017年5月15日 下午4:47:39
	 * @return XSSFSheet[]
	 */
	private XSSFSheet[] createTitles(String[] excleNameArray, XSSFWorkbook wb, String[] colums) {
		// 构建 XSSFSheet 数组
		XSSFSheet[] sheetArray = new XSSFSheet[excleNameArray.length];
		// 创建多标签表头
		for(int i=0; i<excleNameArray.length; i++){
			// 标签名
			String excleName = excleNameArray[i];
			XSSFFont font = wb.createFont();
			font.setColor(XSSFFont.COLOR_NORMAL);
			font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
			XSSFCellStyle titleStyle = wb.createCellStyle();
			titleStyle.setFont(font);
			titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			titleStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
			XSSFSheet sheet = wb.createSheet(excleName);
			XSSFRow titleRow = sheet.createRow(0);
			for (int k = 0; k < colums.length; k++) {
				sheet.setColumnWidth(k, 8000);			
				XSSFCell cell = titleRow.createCell(k, 0);
				cell.setCellStyle(titleStyle);
				cell.setCellType(XSSFCell.CELL_TYPE_STRING);
				cell.setCellValue(colums[k]);
			}
			// 构建标签
			sheetArray[i] = sheet;
		}
		return sheetArray;
	}
	
	/**
	 * 获取Excel所需数组，解析多个数据集合SHEET数据
	 * @author WanTD
	 * @version 2017年5月16日 上午9:19:54
	 * @return List<List<String[]>>
	 */
	@SuppressWarnings("unchecked")
	private List<List<String[]>> formDataColums(List<List<Map<String, Object>>> listdata,String[] colums) throws Exception {
		// Excel所需数组集合
		List<List<String[]>> listArry = new ArrayList<List<String[]>>();
		// 遍历每个数据集合SHEET数据
		for(int i=0;i<listdata.size();i++){
			// 获取每个集合SHEET数据
			List<Map<String, Object>> l = (List<Map<String, Object>>) listdata.get(i);
			List<String[]> list=new ArrayList<>();
			JsonConfig jfg=new JsonConfig();
			jfg.setCycleDetectionStrategy(CycleDetectionStrategy.LENIENT);
			JSONArray jsonArray = JSONArray.fromObject(l, jfg);
			List<JSONObject> mapListJson = jsonArray;
			// 将每个每个集合SHEET数据转换为Excel所需数组集合
			for (JSONObject jobj : mapListJson) {
				// 解析 JsonString,将字段转换为骆驼命名法
				jobj = BeanMapConvertUtil.jsonStrToCmCaseObj(jobj.toString());
				String[] columValue=new String[colums.length];
				for (int j=0;j<colums.length;j++) {
					Object object = jobj.get(colums[j]);
					columValue[j]=object==null?"":object.toString().equals("null")?"":object.toString();
				}
				list.add(columValue);
			}
			listArry.add(list);
		}
		return listArry;
	}
	
	/**
	 * 将多个SHEET标签填充数据
	 * @author WanTD
	 * @version 2017年5月16日 上午9:16:34
	 * @return void
	 */
	private void addDataToSheetArray(XSSFWorkbook workbook, List<List<String[]>> formatelist, XSSFSheet[] sheetArray) {
		// 遍历每个SHEET将数据填充
		for(int x=0;x<sheetArray.length;x++){
			XSSFSheet sheet = sheetArray[x];
			XSSFCellStyle shortStyle = workbook.createCellStyle();
			shortStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);  
			shortStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
			XSSFCellStyle longStyle = workbook.createCellStyle();
			longStyle.setWrapText(true);//自动换行 
			longStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
			longStyle.setAlignment(XSSFCellStyle.ALIGN_LEFT);
			// 获取所需填充数据集合
			List<String[]> list = formatelist.get(x);
			// 将数据填充
			for (String[] strarry:list) {
				XSSFRow row = sheet.createRow(sheet.getLastRowNum()+1);
				for (int i=0;i<strarry.length;i++) {
					creatSheetArrayCell(shortStyle, strarry[i], row, i);
				}
			}
		}
	}

	/**
	 * TODO 业务暂挂
	 * @author WanTD
	 * @version 2017年5月16日 上午9:07:12
	 * @return List<String[]>
	private List<String[]> formDataArray(List listdata,String[] colums) {
		JsonConfig jfg=new JsonConfig();
		jfg.setCycleDetectionStrategy(CycleDetectionStrategy.LENIENT);
		List<String[]> list=new ArrayList<>();
		for(int i=0;i<listdata.size();i++){
			List<Map<String, Object>> l = (List<Map<String, Object>>) listdata.get(0);
			JSONArray jsonArray = JSONArray.fromObject(l, jfg);
			List<JSONObject> mapListJson = (List)jsonArray;
			for (JSONObject jobj : mapListJson) {
				String[] columValue=new String[colums.length];
				for (int j=0;j<colums.length;j++) {
					Object object = jobj.get(colums[j]);
					columValue[j]=object==null?"":object.toString();
				}
				list.add(columValue);
			}
		}
		return list;
	}
	*/

}
