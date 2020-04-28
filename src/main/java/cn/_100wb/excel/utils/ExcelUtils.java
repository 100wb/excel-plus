package cn._100wb.excel.utils;

import org.apache.commons.codec.binary.Base64;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;

/**
 * @description: Excel 工具类
 * @author: WanTD
 * @create: 2020-04-29 01:40
 */
public class ExcelUtils {

    //excel表格导出
    protected  void flushBufferrExcel(XSSFWorkbook wb,String fileName, HttpServletRequest request, HttpServletResponse response) {
//        HttpServletResponse response = response;
        try {
            response.setHeader("Content-Type", "application/vnd.ms-excel;charset=UTF-8");
            response.setHeader("Content-Disposition", "attachment;filename=" + encodeFileName(request, fileName));
            ServletOutputStream outputStream = response.getOutputStream();
            wb.write(response.getOutputStream());
            response.flushBuffer();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 文件名编码
     * @param request
     * @param fileName
     * @return
     * @throws UnsupportedEncodingException
     */
    protected  String encodeFileName(HttpServletRequest request,
                                     String fileName) throws UnsupportedEncodingException {
        String agent = request.getHeader("USER-AGENT");
        if (null != agent && -1 != agent.indexOf("MSIE")) {
            return URLEncoder.encode(fileName, "UTF-8");
        } else if (null != agent && -1 != agent.indexOf("Mozilla")) {
            return "=?UTF-8?B?" + (new String(Base64.encodeBase64(fileName.getBytes("UTF-8"))))+"?=";
        } else {
            return fileName;
        }
    }

}
