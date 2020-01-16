package com.safeschool.admin.utils;


import com.safeschool.admin.annotation.PrintingColumn;
import com.safeschool.admin.annotation.PrintingTitle;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.CollectionUtils;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

/**
 * @author jiangdong
 * @title: ExcelUtil
 * @projectName safe-school-admin
 * @description: TODO
 * @date 2019/10/29 002915:09
 */
@Slf4j
public class ExcelUtil<T> {

    private String queryCriteria;

    public ExcelUtil() {
        this.queryCriteria = null;
    }

    public ExcelUtil(String queryCriteria) {
        this.queryCriteria = queryCriteria;
    }

    public void buildExcel(HttpServletResponse response, List<T> list) throws IOException, IllegalAccessException {
        List<String> parameter = new ArrayList<>();
        String name = "";
        List<Field> fieldArrayList = new ArrayList<>();
        if (!CollectionUtils.isEmpty(list)) {
            Class clazz = list.get(0).getClass();
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                field.setAccessible(true);
                if (field.getAnnotation(PrintingColumn.class) != null) {
                    if (!StringUtils.isEmpty(field.getAnnotation(PrintingColumn.class).name())) {
                        parameter.add(field.getAnnotation(PrintingColumn.class).name());
                    } else {
                        parameter.add(field.getName());
                    }
                    fieldArrayList.add(field);
                }
            }
            PrintingTitle printingTitle = (PrintingTitle) clazz.getDeclaredAnnotation(PrintingTitle.class);
            if (printingTitle != null) {
                name = printingTitle.titleName();
            }
        } else {
            return;
        }
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        final int sheetNum = (int) Math.ceil((float) list.size() / 50000);
        HSSFCellStyle style = hssfWorkbook.createCellStyle();
        style.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        HSSFFont font = hssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        style.setFont(font);
        HSSFCellStyle style2 = hssfWorkbook.createCellStyle();
        style2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        for (int n = 1; n <= sheetNum; n++) {
            final HSSFSheet sheet = hssfWorkbook.createSheet("sheet" + "-" + n);
            List<T> toOut = new ArrayList<>();
            if (sheetNum > 1) {
                if (n == sheetNum) {
                    toOut = getSubList(list, 0, list.size() - 1);
                } else {
                    toOut = getSubList(list, 0, 50000);
                }
            } else {
                toOut = list;
            }
            if (!CollectionUtils.isEmpty(toOut)) {
                Class clazz = toOut.get(0).getClass();
                HSSFRow row1 = sheet.createRow(0);
                HSSFCell cellTitle = row1.createCell(0);
                cellTitle.setCellStyle(style);
                PrintingTitle printingTitle = (PrintingTitle) clazz.getDeclaredAnnotation(PrintingTitle.class);
                if (printingTitle != null) {
                    if (StringUtils.isNotBlank(queryCriteria)) {
                        cellTitle.setCellValue(printingTitle.titleName() + "                         " + queryCriteria);
                    } else {
                        cellTitle.setCellValue(printingTitle.titleName());
                    }
                    sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, printingTitle.totalColumn()));
                }
                sheet.setColumnWidth(0, 10000);
                sheet.setColumnWidth(1, 15000);
                sheet.setColumnWidth(2, 10000);
                sheet.setColumnWidth(3, 10000);
                sheet.setColumnWidth(4, 10000);
                sheet.setColumnWidth(5, 10000);
                sheet.setColumnWidth(6, 10000);
                sheet.setColumnWidth(7, 10000);
                HSSFRow row2 = sheet.createRow(1);
                for (int i = 0; i < parameter.size(); i++) {
                    HSSFCell cell = row2.createCell(i);
                    cell.setCellStyle(style);
                    cell.setCellValue(parameter.get(i));
                }
                for (int i = 0; i < toOut.size(); i++) {
                    HSSFRow row = sheet.createRow(i + 2);
                    for (int j = 0; j < fieldArrayList.size(); j++) {
                        Object value = fieldArrayList.get(j).get(toOut.get(i));
                        HSSFCell cell = row.createCell(j);
                        cell.setCellStyle(style2);
                        if (value != null) {
                            cell.setCellValue(value.toString());
                        } else {
                            cell.setCellValue("");
                        }
                    }
                }
            }
        }
        OutputStream output = response.getOutputStream();
        response.reset();
        name = URLEncoder.encode(name, "UTF-8");
        response.setHeader("Content-disposition", "attachment; filename=" + name + ".xls");
        response.setContentType("application/msexcel");
        hssfWorkbook.write(output);
        output.close();
        log.info("导出excel成功");
    }

    /**截取list  含左不含右
     * @param list
     * @param fromIndex
     * @param toIndex
     * @param <T>
     * @return
     */
    private static <T> List<T> getSubList(List<T> list, int fromIndex, int toIndex) {
        List<T> listClone = list;
        List<T> sub = listClone.subList(fromIndex, toIndex);
        List<T>  two = new ArrayList(sub);
        sub.clear();
        return two;
    }

}
