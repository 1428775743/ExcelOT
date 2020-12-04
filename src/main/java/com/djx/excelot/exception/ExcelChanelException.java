package com.djx.excelot.exception;

public class ExcelChanelException extends Exception{

    private int rowIndex;

    private int cellIndex;

    public ExcelChanelException(int rowIndex, int cellIndex) {
        this.rowIndex = rowIndex;
        this.cellIndex = cellIndex;
    }

    public ExcelChanelException(String message, int rowIndex, int cellIndex) {
        super(message);
        this.rowIndex = rowIndex;
        this.cellIndex = cellIndex;
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
