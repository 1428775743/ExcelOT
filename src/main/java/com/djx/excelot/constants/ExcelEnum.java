package com.djx.excelot.constants;

public enum ExcelEnum {

    V2003(".xls"), V2007(".xlsx");

    private final String suffix;

    private ExcelEnum(String suffix) {
        this.suffix = suffix;
    }

    public String getSuffix() {
        return suffix;
    }

}
