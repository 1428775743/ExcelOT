package com.djx.excelot.exception;

public class ExcelOutLenException extends Exception{

    private int rowIndex;

    private int cellIndex;

    private int maxLen;

    public ExcelOutLenException(int rowIndex, int cellIndex, int maxLen) {
        this.rowIndex = rowIndex;
        this.cellIndex = cellIndex;
        this.maxLen = maxLen;
    }

    public ExcelOutLenException(String message, int rowIndex, int cellIndex, int maxLen) {
        super(message);
        this.rowIndex = rowIndex;
        this.cellIndex = cellIndex;
        this.maxLen = maxLen;
    }

    public int getMaxLen() {
        return maxLen;
    }

    public void setMaxLen(int maxLen) {
        this.maxLen = maxLen;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public int getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(int cellIndex) {
        this.cellIndex = cellIndex;
    }
}
