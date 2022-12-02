package com.carroll.office;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Excel 导出工具类
 *
 * @author: carroll.he
 * @date 2020/5/27
 */
@Slf4j
public class ExportExcelUtils {

    private ExportExcelUtils() {
    }

    /**
     * 填充表头
     * @param headerTitle   表格标题
     * @param sheet         表
     * @param colNm         列标题
     * @param colWidth      列宽
     * @return              返回下一行的行索引
     */
    public static int fillTableHeader(String headerTitle, XSSFSheet sheet, String[] colNm, Integer[] colWidth) {
        int writeCol = 0;
        XSSFCellStyle cellStyleTitle = cellStyleTitle(sheet.getWorkbook(), StyleCategory.HEADER);
        XSSFCellStyle cellStyleColNm = cellStyleTitle(sheet.getWorkbook(), StyleCategory.COLUMN_HEADER);
        cellStyleTitle.setAlignment(HorizontalAlignment.CENTER);
        cellStyleColNm.setAlignment(HorizontalAlignment.CENTER);
        setCellBorder(cellStyleColNm);
        setCellBorder(cellStyleTitle);
        return fillTableHeader(headerTitle, sheet, colNm, colWidth, writeCol, cellStyleTitle, cellStyleColNm);
    }

    /**
     * 填充表头
     * @param headerTitle       表格标题
     * @param sheet             表
     * @param colNm             列标题
     * @param colWidth          列宽
     * @param writeCol          起始列索引
     * @param cellStyleTitle    表标题样式
     * @param cellStyleColNm    列标题样式
     * @return                  返回下一行的行索引
     */
    public static int fillTableHeader(String headerTitle, XSSFSheet sheet, String[] colNm, Integer[] colWidth, int writeCol, XSSFCellStyle cellStyleTitle, XSSFCellStyle cellStyleColNm) {
        int colLength = colWidth.length;
        XSSFRow row = null;
        XSSFCell cell2 = null;
        if (!isNullOrEmpty(headerTitle)) {
            sheet.addMergedRegion(new CellRangeAddress(writeCol, writeCol, 0, colLength - 1));
            row = sheet.createRow(writeCol++);
            row.setHeight((short) 600);
            //表头
            cell2 = row.createCell(0);
            for (int i = 0; i < colLength; i++) {
                cell2 = row.createCell(i);
                cell2.setCellStyle(cellStyleTitle);
                if (i == 0) {
                    cell2.setCellValue(headerTitle);
                }
            }
        }
        row = sheet.createRow(writeCol++);
        row.setHeight((short) 380);
        //列名
        for (int i = 0; i < colLength; i++) {
            sheet.setColumnWidth(i, colWidth[i] * 512);//设置列宽
            XSSFCell cell = row.createCell(i);
            cell.setCellStyle(cellStyleColNm);
            cell.setCellValue(colNm[i]);
        }
        return writeCol;
    }

    /**
     * 设置单元格边框
     */
    public static void setCellBorder(XSSFCellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
    }

    /**
     * 获取样式
     * @param wb
     * @param category  样式类型
     * @return          样式
     */
    public static XSSFCellStyle cellStyleTitle(XSSFWorkbook wb, StyleCategory category) {
        //表头样式
        XSSFCellStyle cellStyle = wb.createCellStyle();
        XSSFFont font = wb.createFont();
        switch (category) {
            case HEADER:
                font.setFontHeightInPoints((short) 20);
                font.setBoldweight(Font.BOLDWEIGHT_BOLD);
                cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
                cellStyle.setFont(font);
                break;
            case MAIN:
                font.setFontHeightInPoints((short) 10);
//      mainfont.setFontName("楷体");
                cellStyle.setFont(font);
                break;
            case COLUMN_HEADER:
                font.setFontHeightInPoints((short) 10);
                font.setBoldweight(Font.BOLDWEIGHT_BOLD);

                cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
                cellStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
                cellStyle.setFont(font);
                break;
            case NO_BORDER:
                font.setFontHeightInPoints((short) 12);
//        fontStyle.setBoldweight(Font.BOLDWEIGHT_BOLD);
                cellStyle.setFont(font);
                break;
        }
        return cellStyle;
    }

    public enum StyleCategory {
        HEADER("表头样式"),
        MAIN("正文样式"),
        COLUMN_HEADER("列名样式"),
        NO_BORDER("无边框，字体大小11");

        private String desc;

        private StyleCategory(String desc) {
            this.desc = desc;
        }

        public String getDesc() {
            return desc;
        }
    }

    /**
     * 填充表数据
     * @param sheet     表
     * @param colDatas  数据
     * @param colWidth  列宽
     * @param writeCol  起始列索引
     */
    public static void fillRowData(XSSFSheet sheet, Object[] colDatas, Integer[] colWidth, int writeCol) {
        XSSFRow row = sheet.createRow(writeCol++);
        //设置备注的样式
        XSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        XSSFFont font = sheet.getWorkbook().createFont();
        font.setFontName("楷体");
        font.setFontHeightInPoints((short) 12);
        cellStyle.setFont(font);
        //添加边框
        setCellBorder(cellStyle);
        row.setHeight((short) 380);
        int colLength = colWidth.length;
        for (int i = 0; i < colLength; i++) {
            //设置列宽
            sheet.setColumnWidth(i, colWidth[i] * 512);
            XSSFCell cell = row.createCell(i);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(colDatas[i] == null ? "" : String.valueOf(colDatas[i]));
        }
    }

    /**
     * 填充数据
     *
     * @param sheet
     * @param rowsData
     * @param colWidth
     * @param writeCol
     */
    public static void fillRowData(XSSFSheet sheet, List<Object[]> rowsData, Integer[] colWidth, int writeCol) {
        int index = 0;
        for (Object[] rowData : rowsData) {
            fillRowData(sheet, rowData, colWidth, writeCol + index);
            index++;
        }
    }

    /**
     * <p> 填写主体</p>
     *
     * @param wb        工作薄
     * @param sheet     工作表
     * @param addList
     * @param objAddr   填写的实体类
     * @param map       当前行数和序号
     * @param fields    属性
     * @param typeNm    分类
     * @param remark    备注换行
     * @param colWidths 列宽
     * @return
     * @throws Exception
     */

    public static Map<String, Integer> addCols(XSSFWorkbook wb, XSSFSheet sheet, List<?> addList, String objAddr, Map<String, Integer> map, String[] fields, String typeNm, String remark, Integer[] colWidths, XSSFCellStyle cellStyleMain) throws Exception {
        int rownumStart = 0;
        int seqNum = 0;
        //设置备注的样式
        XSSFCellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(CellStyle.ALIGN_LEFT);
        XSSFFont font = wb.createFont();
        font.setFontName("楷体");
        font.setFontHeightInPoints((short) 12);
        cellStyle.setFont(font);
        cellStyleMain.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleMain.setWrapText(true);
        //添加边框
        setCellBorder(cellStyle);
        setCellBorder(cellStyleMain);
        XSSFRow row = null;
        XSSFCell cell1 = null;
        int colLength = colWidths.length;
        int arrayLength = colLength;
        if (!isNullOrEmpty(remark)) {
            arrayLength = colLength + 1;
        }

        seqNum = map.get("seqNum");
        rownumStart = map.get("rownumStart");
        if (!isNullOrEmpty(typeNm)) {
            sheet.addMergedRegion(new CellRangeAddress(rownumStart, rownumStart, 0, colLength - 1));
            row = sheet.createRow(rownumStart++);
            row.setHeight((short) 360);
            for (int i = 0; i < colLength; i++) {
                cell1 = row.createCell(i);
                cell1.setCellStyle(cellStyle);
                if (i == 0) {
                    cell1.setCellValue(typeNm);
                }
            }

        }
        try {
            Class<?> c = Class.forName(objAddr);

            for (Object rowObj : addList) {
                row = sheet.createRow(rownumStart++);
                row.setHeight((short) 360);
                for (int i = 0; i < arrayLength; i++) {
                    cell1 = row.createCell(i);
                    if (i < colLength) {
                        cell1.setCellStyle(cellStyleMain);
                    }
                    if (i == 0) {
                        cell1.setCellValue(seqNum++);
                    } else {
                        String fieldVal = "";
                        if (fields[i].startsWith("get")) {
                            Method m = c.getMethod(fields[i], null);
                            Object obj = m.invoke(rowObj, null);
                            fieldVal = obj == null ? "" : (obj + "");
                            //判断“备注”是否为空，不为空换行
                            if (!isNullOrEmpty(remark)) {

                                if (remark.equals(fields[i])) {
                                    if (fieldVal != null && !fieldVal.equals("")) {
                                        sheet.addMergedRegion(new CellRangeAddress(rownumStart, rownumStart, 1, colLength - 1));
                                        row = sheet.createRow(rownumStart++);
                                        cellStyleMain.setWrapText(true);//设置自动换行
                                        cellStyle.setAlignment(CellStyle.VERTICAL_TOP);
                                        cellStyle.setWrapText(true);

                                        float hieght = getExcelCellAutoHeight(fieldVal, 80f);
                                        //根据字符串的长度设置高度
                                        sheet.getRow(sheet.getLastRowNum()).setHeightInPoints(hieght);

                                        for (int b = 0; b < colLength; b++) {
                                            cell1 = row.createCell(b);
                                            cell1.setCellStyle(cellStyle);
                                            if (b == 1) {

                                                cell1.setCellValue(fieldVal);
                                            }
                                        }

                                    }
                                } else {
                                    cell1.setCellValue(fieldVal);
                                    continue;
                                }
                            } else {
                                cell1.setCellValue(fieldVal);
                                continue;
                            }
                        } else if (!isNullOrEmpty(fields[i])) {
                            Object objVal = null;
                            Field field = null;
                            try {
                                field = c.getDeclaredField(fields[i]);
                                field.setAccessible(true);
                                objVal = field.get(rowObj);
                            } catch (Exception e) {
                                log.error(e.getMessage(), e);
                            }
                            if (objVal != null) {
                                if (String.class.equals(field.getType())) {
                                    fieldVal = (String) objVal;
                                } else if (Date.class.equals(field.getType())) {
                                    Date date = (Date) objVal;
                                    fieldVal = getStrDateFormat(date, "yyyy-MM-dd");
                                } else if (java.util.Date.class.equals(field.getType())) {
                                    java.util.Date date = (java.util.Date) objVal;
                                    fieldVal = getStrDateFormat(date, "yyyy-MM-dd");
                                } else if (java.math.BigDecimal.class.equals(field.getType())) {
                                    java.math.BigDecimal bd = (java.math.BigDecimal) objVal;
//							    	fieldVal=bd.toString();
                                    cell1.setCellValue(bd == null ? 0 : bd.doubleValue());
                                    fieldVal = null;
                                } else if (Float.class.equals(field.getType())) {
                                    cell1.getCellStyle().setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                                    cell1.setCellValue(objVal == null ? 0.0 : (Float) objVal);
                                    fieldVal = null;
                                } else if (Double.class.equals(field.getType())) {
                                    cell1.getCellStyle().setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                                    cell1.setCellValue(objVal == null ? 0.0 : (Double) objVal);
                                    fieldVal = null;
                                }
                            }

                        }
                        if (fieldVal != null) {
                            cell1.setCellValue(fieldVal);
                        }
                    }
                }


            }
        } catch (Exception e) {
            throw new Exception("组装下载表格时出错", e);
        }
        map.put("seqNum", seqNum);
        map.put("rownumStart", rownumStart);
        return map;
    }

    public static float getExcelCellAutoHeight(String str, float fontCountInline) {
        //每一行的高度指定
        float defaultRowHeight = 15.00f;
        int defaultCount = 0;
        try {
            defaultCount = str.getBytes("GBK").length;
            defaultCount += (str.split("\\s").length - 1 + str.split("\\n").length - 1) * fontCountInline;
        } catch (UnsupportedEncodingException e) {

        }
        //计算
        return ((int) (defaultCount / (fontCountInline * 2)) + 1) * defaultRowHeight;
    }

    public static void export(String headerTitle, String[] colTitle, Integer[] widthArr, List<Object[]> data, XSSFSheet sheet, OutputStream out) throws Exception {
        int rowIndex = ExportExcelUtils.fillTableHeader(headerTitle, sheet, colTitle, widthArr);
        ExportExcelUtils.fillRowData(sheet, data, widthArr, rowIndex);
        writeExport(sheet.getWorkbook(), out);
    }

    public static void export(String headerTitle, String[] colTitle, Integer[] widthArr, List<Object[]> data, String sheetName, String fileName) throws Exception {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet(sheetName);
        FileOutputStream fout = new FileOutputStream(fileName);
        export(headerTitle, colTitle, widthArr, data, sheet, fout);
    }

    public static void writeExport(XSSFWorkbook wb, OutputStream out) throws Exception {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        wb.write(outputStream);
        byte[] content = outputStream.toByteArray();
        InputStream is = new ByteArrayInputStream(content);
        BufferedInputStream bis = null;
        BufferedOutputStream bos = null;
        try {
            bis = new BufferedInputStream(is);
            bos = new BufferedOutputStream(out);
            byte[] buff = new byte[2048];
            int bytesRead;
            // Simple read/write loop.
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
        } catch (final IOException e) {
            throw e;
        } finally {
            if (bis != null) {
                bis.close();
            }
            if (bos != null) {
                bos.close();
            }
        }
    }

    private static String getStrDateFormat(java.util.Date date, String format) {
        SimpleDateFormat dateFormat = new SimpleDateFormat(format);
        return null != date ? dateFormat.format(date) : "";
    }

    private static boolean isNullOrEmpty(String str) {
        return str == null || "".equals(str);
    }

    public static void main(String[] args) throws Exception {
//        XSSFWorkbook wb = new XSSFWorkbook();
//        XSSFSheet sheet = wb.createSheet("test");
        String[] colNames = new String[]{"test1", "test2 xxxx"};
        Integer[] colWidth = new Integer[]{5, 10};
        List<Object[]> data = new ArrayList<>();
        data.add(new String[]{"0123", "xxxxx"});
        export("test header title", colNames, colWidth, data, "test sheet", "d:/test.xlsx");
//        int rowIndex = fillTableHeader("", sheet, colNames, colWidth);
//        fillRowData(sheet, new String[]{"0123", "xxxxx"}, colWidth, rowIndex);
//        rowIndex++;
//        fillRowData(sheet, new String[]{"01234", "0xxxxx"}, colWidth, rowIndex);
//
//        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
//        wb.write(outputStream);
//        byte[] content = outputStream.toByteArray();
//        InputStream is = new ByteArrayInputStream(content);
//        // 设置response参数，可以打开下载页面
//        FileOutputStream out = new FileOutputStream("D:/test.xls");
//        BufferedInputStream bis = null;
//        BufferedOutputStream bos = null;
//        try {
//            bis = new BufferedInputStream(is);
//            bos = new BufferedOutputStream(out);
//            byte[] buff = new byte[2048];
//            int bytesRead;
//            // Simple read/write loop.
//            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
//                bos.write(buff, 0, bytesRead);
//            }
//        } catch (final IOException e) {
//            throw e;
//        } finally {
//            if (bis != null) {
//                bis.close();
//            }
//            if (bos != null) {
//                bos.close();
//            }
//        }
    }
}
