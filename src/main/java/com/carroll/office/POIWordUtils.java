package com.carroll.office;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.math.BigInteger;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author: carroll.he
 * @date 2020/6/15
 */
@Slf4j
public class POIWordUtils {
    public static XWPFTable addTable(CustomXWPFDocument document, List<Map<String, Object>> data, List<String> headerKeys, List<String> headers, int[] colWidths) {
        return addTable(document, data, headerKeys, headers, colWidths, null, null);
    }

    public static XWPFTable addTable(CustomXWPFDocument document, List<Map<String, Object>> data, List<String> headerKeys, List<String> headers, int[] colWidths, Style headerStyle, Style bodyStyle) {
        Map<String, Object> header = new HashMap<>();
        for (int i = 0; i < headers.size() && i < headerKeys.size(); i++) {
            header.put(headerKeys.get(i), headers.get(i));
        }
        data.add(0, header);
        XWPFTable table = document.createTable();
        XWPFTableRow row = null;
        XWPFTableCell cell = null;
        int index = 0;
        int cellIndex = 0;
        XWPFRun run = null;
        String text = null;
        Style style = null;
        for (Map<String, Object> rowData : data) {
            if (index == 0) {
                row = table.getRow(0);
            } else {
                row = table.createRow();
            }
            cellIndex = 0;
            for (String key : headerKeys) {
                text = rowData.get(key) != null ? String.valueOf(rowData.get(key)) : "";
                style = bodyStyle;
                if (index == 0) {
                    if (cellIndex != 0) {
                        cell = row.addNewTableCell();
                    } else {
                        cell = row.getCell(cellIndex);
                    }
                    style = headerStyle;
                } else {
                    cell = row.getCell(cellIndex);
                }
                setText(cell.addParagraph(), text, style);
                cellIndex++;
            }
            index++;
        }
        setTableGridCol(table, colWidths);
        return table;
    }

    public static void setText(XWPFParagraph paragraph, String text, Style style) {
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        if (style != null) {
            if (style.getAlignment() != null) {
                paragraph.setAlignment(style.getAlignment());
            }
            if (style.getFontFamily() != null) {
                run.setFontFamily(style.getFontFamily());
            }
            if (style.getFontSize() > 0) {
                run.setFontSize(style.getFontSize());
            }
            run.setBold(style.isBold());
        }
    }

    /**
     * @Description: 设置表格列宽
     */
    public static void setTableGridCol(XWPFTable table, int[] colWidths) {
        int index = 0;
        for (XWPFTableRow row : table.getRows()) {
            index = 0;
            for (XWPFTableCell cell : row.getTableCells()) {
                CTTcPr cellPr = cell.getCTTc().addNewTcPr();
                CTTblWidth tblWidth = cellPr.isSetTcW() ? cellPr.getTcW() : cellPr.addNewTcW();
                tblWidth.setType(STTblWidth.DXA);
                tblWidth.setW(new BigInteger(String.valueOf(colWidths[index])));
                index++;
            }
        }

//        CTTbl ttbl = table.getCTTbl();
//        CTTblGrid tblGrid = ttbl.getTblGrid() != null ? ttbl.getTblGrid()
//                : ttbl.addNewTblGrid();
//        for (int j = 0, len = colWidths.length; j < len; j++) {
//            CTTblGridCol gridCol = tblGrid.addNewGridCol();
//            gridCol.setW(new BigInteger(String.valueOf(colWidths[j])));
//        }
    }

    public static void addPic(CustomXWPFDocument document, String filePath, int width, int height, String picAttch) throws Exception {
        XWPFParagraph paragraph = document.createParagraph();
        FileInputStream in = new FileInputStream(filePath);
        String ind = document.addPictureData(in, XWPFDocument.PICTURE_TYPE_JPEG);
        System.out.println("pic ID=" + ind);
        document.createPicture(paragraph, document.getAllPictures().size() - 1, width, height, picAttch);
    }

    public static void addPic(CustomXWPFDocument document, File file, String picAttch) throws Exception {
        XWPFParagraph paragraph = document.createParagraph();
        FileInputStream in = new FileInputStream(file);
        BufferedImage img = javax.imageio.ImageIO.read(in);
        int width = img.getWidth();
        int height = img.getHeight();
        int picWidth = 600;
        int picHeight = picWidth * height / width;
        in = new FileInputStream(file);
        String ind = document.addPictureData(in, XWPFDocument.PICTURE_TYPE_JPEG);
        log.debug("pic ID={},height={},width={}", ind, picHeight, picWidth);
//        document.createPicture(paragraph, document.getAllPictures().size() - 1, Units.toEMU(picWidth), Units.toEMU(picHeight), picAttch);
        document.createPicture(paragraph, document.getAllPictures().size() - 1, picWidth, picHeight, picAttch);
    }

    public static XWPFRun addTitle(CustomXWPFDocument document, String style, String title) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setStyle(style);
        XWPFRun run2 = paragraph.createRun();
        run2.setText(title);
        return run2;
    }

    public static void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel) {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);

        // is a null op if already defined
        XWPFStyles styles = docxDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);
    }
}
