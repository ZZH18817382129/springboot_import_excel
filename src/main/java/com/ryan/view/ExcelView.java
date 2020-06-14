package com.ryan.view;

import com.ryan.entry.ExcelSheetSettingEnum;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.http.HttpHeaders;
import org.springframework.web.servlet.view.document.AbstractXlsView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @Author: Administrator
 * @Date: 2020/5/28 0:21
 */
public class ExcelView extends AbstractXlsView {
    @Override
    protected void buildExcelDocument(Map<String, Object> map, Workbook workbook, HttpServletRequest httpServletRequest, HttpServletResponse httpServletResponse) throws Exception {
        ExcelSheetSettingEnum setting = (ExcelSheetSettingEnum) map.get("ExcelSheetSetting");
        //设置文件名称
        String fliename = setting.getFilename();
        //设置编码
        fliename = new String(fliename.getBytes("UTF-8"), "ISO8859-1");
        //输出格式
        httpServletResponse.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment;fliename=" + fliename + ".xls");
        List<List<List<String>>> sheets = (List<List<List<String>>>) map.get("date");
        for (int i = 0; i < sheets.size(); i++) {
            //创建sheet
            String[] sheetNames = setting.getSheetnames();
            String sheetName = "Sheet" + (i + 1);
            if (sheetNames != null && sheetNames.length > 0) {
                sheetName = sheetNames[i];
            }
            Sheet sheet = workbook.createSheet(sheetName);
            //如果标题不为空的话，将表格的第一列作为标题行，并合并第一行的N个单元格
            int index = 0;
            String[] titles = setting.getTitles();
            String[][] headers = setting.getHeaders();
            List<List<String>> rowsForTable = sheets.get(i);
            if (titles != null && titles.length > 0) {
                // 合并标题单元格 下标从0开始 起始行号，终止行号， 起始列号，终止列号
                CellRangeAddress region = new CellRangeAddress(0, 0, 0, rowsForTable.get(0).size() - 1);
                sheet.addMergedRegion(region);
                Row titleRow = sheet.createRow(index++);
                Cell titleCell = titleRow.createCell(0);
                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
                Font font = workbook.createFont();
                font.setFontName("黑体");
                font.setBold(true);
                font.setFontHeightInPoints((short) 15);
                cellStyle.setFont(font);
                titleCell.setCellStyle(cellStyle);
                titleCell.setCellValue(titles[i]);
            }
            // 创建表头行
            if (headers != null && headers.length > 0) {
                Row headerRow = sheet.createRow(index++);
                String[] header = headers[i];
                for (int j = 0; j < headers.length; j++) {
                    headerRow.createCell(j).setCellValue(header[j]);
                }
            }
            // 创建数据行
            AtomicInteger rowIndex = new AtomicInteger(index);
            rowsForTable.forEach(rowList -> {
                Row row = sheet.createRow(rowIndex.getAndIncrement());
                AtomicInteger x = new AtomicInteger();
                rowList.forEach(cell ->
                        row.createCell(x.getAndIncrement()).setCellValue(cell)
                );
            });
        }
    }
}
