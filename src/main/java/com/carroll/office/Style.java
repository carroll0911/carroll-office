package com.carroll.office;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;

/**
 * @author: carroll.he
 * @date 2020/6/18
 * Copyright @2020 Tima Networks Inc. All Rights Reserved. 
 */
public class Style {

    private String fontFamily;

    private int fontSize;

    private boolean bold;

    private ParagraphAlignment alignment;

    public Style() {
    }

    public Style(String fontFamily, int fontSize, boolean bold, ParagraphAlignment alignment) {
        this.fontFamily = fontFamily;
        this.fontSize = fontSize;
        this.bold = bold;
        this.alignment = alignment;
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public void setFontFamily(String fontFamily) {
        this.fontFamily = fontFamily;
    }

    public int getFontSize() {
        return fontSize;
    }

    public void setFontSize(int fontSize) {
        this.fontSize = fontSize;
    }

    public boolean isBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public ParagraphAlignment getAlignment() {
        return alignment;
    }

    public void setAlignment(ParagraphAlignment alignment) {
        this.alignment = alignment;
    }
}
