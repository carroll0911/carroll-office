package com.carroll.office;

/**
 * @author: carroll.he
 * @date 2020/5/26
 * Copyright @2020 Tima Networks Inc. All Rights Reserved. 
 */
public class OfficeException extends Exception {
    private String code;

    public OfficeException(String code, String msg) {
        super(msg);
        this.code = code;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }
}
