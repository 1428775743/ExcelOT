package com.djx.excelot.entity;

import com.djx.excelot.annocation.CellDouble;
import com.djx.excelot.annocation.CellValue;
import com.djx.excelot.annocation.Excel;

@Excel(
        name = "老师",
        sheetName = "sheet1"
)
public class Teacher {

    @CellValue(name = "名字",index = 1)
    String teachername;

    @CellDouble(name = "收入",index = 2, fixed = 2)
    private Double balance;

    public String getTeachername() {
        return teachername;
    }

    public void setTeachername(String teachername) {
        this.teachername = teachername;
    }

    public Double getBalance() {
        return balance;
    }

    public void setBalance(Double balance) {
        this.balance = balance;
    }

    @Override
    public String toString() {
        return "Teacher{" +
                "teachername='" + teachername + '\'' +
                ", balance=" + balance +
                '}';
    }
}
