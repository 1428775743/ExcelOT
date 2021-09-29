package com.djx.excelot.exception;

public class ExcelDateParseException extends Exception{

    private int rowIndex;

    private int cellIndex;

    public ExcelDateParseException(int rowIndex, int cellIndex) {
        this.rowIndex = rowIndex;
        this.cellIndex = cellIndex;
    }

    public ExcelDateParseException(String message, int rowIndex, int cellIndex) {
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
