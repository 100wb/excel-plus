package cn._100wb.excel.utils;

import java.lang.reflect.Field;
import java.util.Date;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import net.sf.cglib.beans.BeanMap;
import net.sf.json.JSONObject;

/**
 * BeanMap 转换工具类
 * @author WanTD
 * @version 2018年2月9日上午8:56:20
 */
public class BeanMapConvertUtil {
	
	//  Date Column 分隔符
	public static String FORMATCOLUMNSPLIT = "#_#";
	
	//  Date Format Code
	private static String DATEFORMATCODE = "yyyy-MM-dd HH:mm:ss";
	
    /**
     * 将 List<JavaBean>对象转化为List<Map>
     * @author WanTD
     * @version 2018年3月27日 上午11:25:10
     * @return List<Map<String,Object>>
     */
    public static <T> List<Map<String, Object>> convertListBeanListMap(List<T> beanList, String formatColumnDate) throws Exception {
        List<Map<String, Object>> mapList = new ArrayList<>();
        for (int i = 0, n = beanList.size(); i < n; i++)
        {
            Object bean = beanList.get(i);
            Map<String, Object> map = beanToMap(bean, formatColumnDate);
            mapList.add(map);
        }
        return mapList;
    }

	/**
	 * Bean 转换 Map
	 * @author WanTD
	 * @param bean 需要转换的 Bean
	 * @param formatDateCode c_array 需要对外键转换的Class数组
	 * @version 2018年3月27日 上午9:15:45
	 * @return Map<String,Object>
	 */
	@SuppressWarnings("unchecked")
	public static <T> Map<String, Object> beanToMap(T bean, String formatDateCode) throws Exception {  
	    Map<String, Object> map = new LinkedHashMap<String, Object>();  
	    if (bean != null){
	    	// 优化解析数据效率,采用net.sf.cglib.beans.BeanMap类中的方法,这种方式效率极高
	        BeanMap beanMap = BeanMap.create(bean); 
	        // 遍历数据Bean每个字段的集合Map数据
	        for (Object key : beanMap.keySet()){
	        	// 是否为空
        		if(null != beanMap.get(key)){
        			// 是否为基础类型
        			if(null == beanMap.get(key).getClass().getClassLoader()){
        				// Declared Field
        				Field declaredField = bean.getClass().getDeclaredField(key.toString());
        				// 扩展 - 时间数据格式化
        				if(Date.class == declaredField.getType()){
        					// Date Format
        					String dateFormat = (null == formatDateCode && !"".equals(formatDateCode)) ? DATEFORMATCODE : formatDateCode;
        					// 构建格式数据格式
							java.text.SimpleDateFormat formatDateColumn = new java.text.SimpleDateFormat(dateFormat);
							// 格式化后 Column Date 数据
							String dateColumn = formatDateColumn.format(formatDateColumn.parse(beanMap.get(key).toString()));
	    					map.put(key+"", dateColumn);
        				}else if(Map.class == declaredField.getType()){
        					// 转换Map集合格式
        					Map<String, Object> map_result = (Map<String, Object>) beanMap.get(key);
        					for(String map_key : map_result.keySet()) {
        						T entity = (T) map_result.get(map_key);
        						// 将外键 Map 数据,二次解析
        						BeanMap beanMapTemp = BeanMap.create(entity);
        						Map<String, Object> mapTemp = new LinkedHashMap<String, Object>();  
        						for (Object keyTemp : beanMapTemp.keySet()) {
        							// 过滤叠加外键
        							if(null != beanMapTemp.get(keyTemp) && null == beanMapTemp.get(keyTemp).getClass().getClassLoader()){
        								mapTemp.put(keyTemp+"", beanMapTemp.get(keyTemp)+"");  
        							}
        						}
        						// 解析后,外键Map数据
        						map.put(map_key+"", mapTemp);
    					    }
        				}else{
        					// 无需解析
        					map.put(key+"", beanMap.get(key));
        				}
        			}else{
        				// 将外键 Map 数据,二次解析
						BeanMap beanMapTemp = BeanMap.create(beanMap.get(key));
						Map<String, Object> mapTemp = new LinkedHashMap<String, Object>();  
						for (Object keyTemp : beanMapTemp.keySet()) {
							mapTemp.put(keyTemp+"", beanMapTemp.get(keyTemp)+"");  
						}
						// 解析后,外键Map数据
						map.put(key+"", mapTemp);
        			}
        		}else{
        			// 字段为空,这里可针对业务进行维护
        			map.put(key+"", "");
        		}
	        }             
	    }  
	    return map;  
	}
	
	/**
	 * 针对Map Key默认大写,为 JsonArray,JsonObject To Key
	 * 优化繁琐拆装箱操作,动态构建 Map Key 将字段转化为骆驼峰命名风格
	 * @author WanTD
	 * @version 2018年12月21日 上午9:06:04
	 * @return JSONObject
	 */
	public static JSONObject jsonStrToCmCaseObj(String jsonObj) throws Exception {
		// 解析 JsonString Map Key 正则
		String regObjMap = "\"(\\w+)\":+?";
		// 正则表达式解析,匹配目标 Map Key
		Matcher matcher = Pattern.compile(regObjMap).matcher(jsonObj);
		// 逐个匹配 Map Key
		while(matcher.find()){
			// 解析 Object Map Key 正则
			String objMapKey = ".+\\_(\\w).+";
			// 解析非骆驼命名法字段转换小写
			if(-1 == matcher.group(1).toString().lastIndexOf("_") || 0 == matcher.group(1).toString().lastIndexOf("_")){
				// 将解析后 keyCmCase 替换
	    		jsonObj = jsonObj.replace(matcher.group(0), "\""+ Pattern.compile(objMapKey).matcher(matcher.group(1).toLowerCase()).replaceAll("$1").toLowerCase() +"\":");
				continue;
			}
			// 固定获取分组数据,因此不需要使用 for
			Matcher matcheCmCase = Pattern.compile("\\_(\\w)+?").matcher(matcher.group(1).toLowerCase());
			// 获取需要解析骆驼命名法字段Key
			String keyCmCase = matcher.group(1).toLowerCase();
			// 按骆驼命名法,循环解析下划线第一位将其转换大写
			while(matcheCmCase.find()){
				keyCmCase = keyCmCase.replaceAll("\\_"+matcheCmCase.group(1).toLowerCase(), matcheCmCase.group(1).toUpperCase());
				// 业务暂挂,后续可优化开放...待定
				// Matcher matcherMapKey = Pattern.compile(objMapKey).matcher(matcher.group(1).toLowerCase());
				// String keyCmCase = matcher.group(1).toLowerCase().replaceAll("\\_(\\w)", matcherMapKey.replaceAll("$1").toUpperCase());
			}
    		// 将解析后 keyCmCase 替换
    		jsonObj = jsonObj.replace(matcher.group(0), "\""+ keyCmCase +"\":");
		}
		// Map Key解析后,转换成 JsonObject
		return JSONObject.fromObject(jsonObj);
	}
	
	/**
	 * 追加SQL字段别名
	 * @param column SQL字段数据
	 * @param clmAlias SQL别名
	 * @author WanTD
	 * @version 2018年12月25日 上午10:58:16
	 * @return String
	 */
	public String appendColumnAlias(String column, String clmAlias) throws Exception {
		// 获取字段是否包含分隔符状态
		boolean clmSltFlag = (-1 != column.trim().indexOf(",")) ? true : false;
		// 解析SQL字段正则
		Pattern r = Pattern.compile("(.+?)\\.(.+)?\\,|(.+?)\\.(.+)?");
		// 正则解析数据匹配
		Matcher m = r.matcher(column);
		// 数据是否符合正则匹配
		if(m.matches()) {
			// 获取[表/字段]别名,
			String tabAS= (null != m.group(1)) ? m.group(1).trim() : m.group(3).trim(), 
				   clmAS = (null != m.group(2)) ? m.group(2).trim() : m.group(4).trim();
			// 解析 Object Map Key 正则
			String objMapKey = ".+\\_(\\w).+";
			// 固定获取分组数据,因此不需要使用 for
	   		Matcher matcherMapKey = Pattern.compile(objMapKey).matcher(clmAS);
	   		// 按骆驼命名法解析下划线第一位将其转换大写
	   		String keyCmCase = clmAS.replaceAll("\\_(\\w)", matcherMapKey.replaceAll("$1").trim().toUpperCase());
				// 是否别名代替默认骆驼命名法
	   		if(null != clmAlias){
					column = tabAS+"."+clmAS+" \""+clmAlias+"\""+(!clmSltFlag ? "" : ", ");
	   		} else {
	   			column = tabAS+"."+clmAS+" \""+keyCmCase+"\""+(!clmSltFlag ? "" : ", ");
	   		}
		}
		return column;
   }

}
