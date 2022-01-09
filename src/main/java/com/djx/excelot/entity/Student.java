package com.djx.excelot.entity;

import com.djx.excelot.annocation.*;
import com.djx.excelot.constants.ExcelEnum;
import lombok.Data;

import java.util.Date;

@Excel(
        name = "学生统计数据表",
        sheetName = "Sheet1"
)
@Data
public class Student {

    public Student() {
    }

    public Student(Long id, String name, Integer age, Date date, Boolean isDel, Double balance, Integer select, String select2) {
        this.id = id;
        this.name = name;
        this.age = age;
        this.date = date;
        this.isDel = isDel;
        this.balance = balance;
        this.select = select;
        this.select2 = select2;
    }

    private Long id;

    /**
     * value 支持绝大多数常用类型 都可以直接显示
     */
    @CellWidth(width = 150)
    @CellValue(name = "名字",index = 0)
    private String name;

    /**
     * value 支持绝大多数常用类型 都可以直接显示
     */
    @CellWidth(width = 50)
    @CellValue(name = "年龄",index = 1)
    private Integer age;

    @CellWidth(width = 250)
    @CellDate(name = "日期",index = 2,formatStr = "yyyy-MM-dd hh:mm:ss")
    private Date date;

    /**
     * boolean类型 需要在excel变成其他值
     */
    @CellWidth(width = 50)
    @CellBoolean(name = "是否删除",index = 3,tureValue = "是", falseValue = "否")
    private Boolean isDel;

    /**
     * double类型 需要保留位数
     */
    @CellWidth(width = 100)
    @CellDouble(name = "收入",index = 4, fixed = 2)
    private Double balance;


    /**
     * 枚举类型 可以使用数字 keys 就是数据值 values 就是对应显示的值
     */
    @CellWidth(width = 100)
    @CellSelect(
            name = "选择",
            index = 5,
            keys = {"1","2","3"},
            values = {"选择1","选择2","选择3"}
    )
    private Integer select;

    /**
     * 枚举类型 可以使用字符串
     */
    @CellWidth(width = 100)
    @CellSelect(
            name = "选择2",
            index = 6,
            keys = {"a","b","c"},
            values = {"2选择1","2选择2","2选择3"}
    )
    private String select2;

    /**
     * Formula 填excel公式进
     * 下面这个就是 计算 e/b 的值 就是计算    下标是4(收入) / 下标是1(年龄) 的值 保留两位小数点
     */
    @CellFormula(name = "年收比", fomula = "=round(E#index/B#index,2)", index = 7)
    private Double testFormula;

}
