package com.djx.excelot.entity;

import com.djx.excelot.annocation.*;
import com.djx.excelot.constants.ExcelEnum;

import java.util.Date;

@Excel(
        name = "学生统计数据表",
        sheetName = "sheet1"
)
public class Student {

    private Long id;

    @CellValue(name = "名字",index = 0)
    private String name;

    @CellValue(name = "年龄",index = 1,suffix = "岁")
    private Integer age;

    @CellDate(name = "日期",index = 2,formatStr = "yyyy-MM-dd hh:mm:ss")
    private Date date;

    @CellBoolean(name = "是否删除",index = 3,tureValue = "是", falseValue = "否")
    private Boolean isDel;

    @CellDouble(name = "收入",index = 4, fixed = 2)
    private Double balance;

    @CellSelect(
            name = "选择",
            index = 5,
            keys = {"1","2","3"},
            values = {"选择1","选择2","选择3"}
    )
    private Integer select;

    @CellSelect(
            name = "选择2",
            index = 6,
            keys = {"a","b","c"},
            values = {"2选择1","2选择2","2选择3"}
    )
    private String select2;

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

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public Boolean getIsDel() {
        return isDel;
    }

    public void setIsDel(Boolean del) {
        isDel = del;
    }

    public Double getBalance() {
        return balance;
    }

    public void setBalance(Double balance) {
        this.balance = balance;
    }

    public Integer getSelect() {
        return select;
    }

    public void setSelect(Integer select) {
        this.select = select;
    }

    public String getSelect2() {
        return select2;
    }

    public void setSelect2(String select2) {
        this.select2 = select2;
    }

    @Override
    public String toString() {
        return "Student{" +
                "id=" + id +
                ", name='" + name + '\'' +
                ", age=" + age +
                ", date=" + date +
                ", isDel=" + isDel +
                ", balance=" + balance +
                ", select=" + select +
                ", select2='" + select2 + '\'' +
                '}';
    }
}
