package com.djx.excelot;

import com.djx.excelot.entity.Student;
import com.djx.excelot.entity.Teacher;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelOTExportTest {

    public static void main(String[] args) throws Exception {

        File file = new File("F:\\student.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        // 使用
        ExcelImportAndExport<Teacher> excelUtils = new ExcelImportAndExport<>();
        List<Teacher> list = excelUtils.importExcel(fileInputStream, "student.xlsx", Teacher.class);

        System.out.println(list);
    }
}
