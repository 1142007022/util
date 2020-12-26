package com.jiangdong.utils;


import cn.hutool.core.util.URLUtil;
import com.jiangdong.annotation.Column;
import com.jiangdong.annotation.Title;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.sql.Time;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author jiangdong
 * @title: ExcelUtil
 * @projectName utils
 * @description: 默认每个sheet最多50000条数据  超过另起一个sheet
 * @date 2019/10/29 002915:09
 */
@Data
public class ExcelUtil<T> {

    /**
     * 设置每行的宽度 每个值的index 对应第几列  如{1500,1000} 表示第一个1500长度 第二个1000长度 以此类推
     */
    private int[] size;

    /**
     * 查询条件文本描述
     */
    private String queryCriteria;

    /**
     * 是否关闭流 默认关闭
     */
    private boolean close = true;

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
        if (list != null && list.size() > 0) {
            Class clazz = list.get(0).getClass();
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                field.setAccessible(true);
                if (field.getAnnotation(Column.class) != null) {
                    if (!StringUtils.isEmpty(field.getAnnotation(Column.class).name())) {
                        parameter.add(field.getAnnotation(Column.class).name());
                    } else {
                        parameter.add(field.getName());
                    }
                    fieldArrayList.add(field);
                }
            }
            Title title = (Title) clazz.getDeclaredAnnotation(Title.class);
            if (title != null) {
                name = title.title();
            }
        } else {
            return;
        }

        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        final int sheetNum = (int) Math.ceil((float) list.size() / 50000);
        HSSFCellStyle style = hssfWorkbook.createCellStyle();
        style.setFillForegroundColor((short) 22);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = hssfWorkbook.createFont();
        font.setFontHeightInPoints((short) 12);
        style.setFont(font);
        HSSFCellStyle style2 = hssfWorkbook.createCellStyle();
        style2.setAlignment(HorizontalAlignment.CENTER);

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
            if (toOut != null && toOut.size() > 0) {
                Class clazz = toOut.get(0).getClass();
                HSSFRow row1 = sheet.createRow(0);
                HSSFCell cellTitle = row1.createCell(0);
                cellTitle.setCellStyle(style);
                Title title = (Title) clazz.getDeclaredAnnotation(Title.class);
                if (title != null) {
                    if (StringUtils.isNotBlank(queryCriteria)) {
                        cellTitle.setCellValue(title.title() + "                         " + queryCriteria);
                    } else {
                        cellTitle.setCellValue(title.title());
                    }
                    sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, parameter.size() - 1));
                }
                if (getSize() != null && getSize().length > 0) {
                    for (int i = 0; i < getSize().length; i++) {
                        sheet.setColumnWidth(i, getSize()[i]);
                    }
                } else {
                    int size = parameter.size();
                    this.size = new int[size];
                    for (int i = 0; i < size; i++) {
                        this.size[i] = 10000;
                        sheet.setColumnWidth(i, getSize()[i]);
                    }
                }
                HSSFRow row2 = sheet.createRow(1);
                for (int i = 0; i < parameter.size(); i++) {
                    HSSFCell cell = row2.createCell(i);
                    cell.setCellStyle(style);
                    cell.setCellValue(parameter.get(i));
                }
                for (int i = 0; i < toOut.size(); i++) {
                    HSSFRow row = sheet.createRow(i + 2);
                    for (int j = 0; j < fieldArrayList.size(); j++) {
                        Field field = fieldArrayList.get(j);
                        Object value = field.get(toOut.get(i));
                        HSSFCell cell = row.createCell(j);
                        cell.setCellStyle(style2);
                        Column column = field.getDeclaredAnnotation(Column.class);
                        if (value != null) {
                            String rule = column.timeFormat();
                            boolean rate = column.rate();
                            if (StringUtils.isNotBlank(rule) && (field.getType().equals(Date.class) ||
                                    field.getType().equals(java.sql.Date.class) ||
                                    field.getType().equals(Time.class) ||
                                    field.getType().equals(Timestamp.class))) {
                                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(rule);
                                cell.setCellValue(simpleDateFormat.format(value));
                            } else if (rate && field.getType().equals(BigDecimal.class)) {
                                BigDecimal valueReal = (BigDecimal) value;
                                cell.setCellValue(valueReal.multiply(new BigDecimal("100")).toString() + "%");
                            } else {
                                cell.setCellValue(value.toString());
                            }

                        } else {
                            if (field.getType().equals(Integer.class) || field.getType().equals(Long.class) ||
                                    field.getType().equals(Double.class) || field.getType().equals(Float.class) ||
                                    field.getType().equals(BigDecimal.class)) {
                                cell.setCellValue(0);
                            } else {
                                cell.setCellValue("");
                            }
                        }
                    }
                }
            }
        }

        OutputStream output = response.getOutputStream();
        response.reset();
        response.setCharacterEncoding("utf-8");
        name = URLEncoder.encode(name + ".xls", "UTF-8");
        //response.setHeader("Content-disposition", "attachment; filename=" + name + ".xls");
        response.setHeader("Cache-Control", "no-cache, no-store, must-revalidate");
        response.setHeader("Content-Disposition", "attachment; filename=" + URLUtil.encode(name, "UTF-8"));
        response.setHeader("Pragma", "no-cache");
        response.setHeader("Expires", "0");
        response.setContentType("application/msexcel;charset=utf-8");
        hssfWorkbook.write(output);
        if (close) {
            output.close();
        }
    }

    /**
     * 截取list  含左不含右
     *
     * @param list
     * @param fromIndex
     * @param toIndex
     * @param <T>
     * @return
     */
    private static <T> List<T> getSubList(List<T> list, int fromIndex, int toIndex) {
        List<T> listClone = list;
        List<T> sub = listClone.subList(fromIndex, toIndex);
        List<T> two = new ArrayList(sub);
        sub.clear();
        return two;
    }
}
