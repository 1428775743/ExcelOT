package com.djx.excelot;

import com.djx.excelot.annocation.*;
import com.djx.excelot.annocation.CellValue;
import com.djx.excelot.constants.ExcelEnum;
import com.djx.excelot.exception.ExcelChanelException;
import com.djx.excelot.exception.ExcelNullpointExcetion;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelImportAndExport<T> {


    private final Excel excelConfig;

    private Workbook workbook;

    private Sheet selectSheet;

    private final Class<T> cls;

    private final List<Excelmode> excelmodeList;

    private int lastIndex = -1;

    public ExcelImportAndExport(Class<T> cls) throws Exception {
        this.cls = cls;
        excelConfig = cls.getAnnotation(Excel.class);

        if (excelConfig == null) {
            throw new Exception(cls.getName() + "没有@Excel注解");
        }

        excelmodeList = this.handleCls(cls);
    }

    /**
     * 将对象 按照注解转换成Excelworkbook
     * @param objs
     * @return
     */
    public Workbook exportExcel(List<T> objs) throws Exception {

        // 配置参数
        ExcelEnum excelVersion = excelConfig.version();
        String sheetName = excelConfig.sheetName();

        Workbook workbook = getWorkbook(excelVersion);
        Sheet sheet = workbook.createSheet(sheetName);

        Row rowhead = sheet.createRow(0);
        // 列名
        for (Excelmode excelmode: excelmodeList) {
            Cell cell = rowhead.createCell(excelmode.index);
            cell.setCellValue(excelmode.name);
        }

        for (int i = 0;i < objs.size(); i++) {
            T obj = objs.get(i);
            Row row = sheet.createRow(i+1);

            for (Excelmode excelmode: excelmodeList) {
                Cell cell = row.createCell(excelmode.index);
                Annotation annotation = excelmode.annotation;

                try {
                    if (annotation.annotationType().equals(CellValue.class)) {
                        cell.setCellValue(cellValue((CellValue) annotation, excelmode.method.invoke(objs.get(i))));
                        continue;
                    }

                    if (annotation.annotationType().equals(CellSelect.class)) {
                        cell.setCellValue(cellSelect((CellSelect) annotation, excelmode.method.invoke(objs.get(i))));
                        continue;
                    }

                    if (annotation.annotationType().equals(CellDouble.class)) {
                        cell.setCellValue(cellDouble((CellDouble) annotation, excelmode.method.invoke(objs.get(i))));
                        continue;
                    }

                    if (annotation.annotationType().equals(CellDate.class)) {
                        cell.setCellValue(cellDate((CellDate) annotation, excelmode.method.invoke(objs.get(i))));
                        continue;
                    }

                    if (annotation.annotationType().equals(CellBoolean.class)) {
                        cell.setCellValue(cellBoolean((CellBoolean) annotation, excelmode.method.invoke(objs.get(i))));
                        continue;
                    }
                } catch (Exception e) {
                    throw new ExcelChanelException(cell.getRowIndex(), cell.getColumnIndex());
                }

            }
        }

        return workbook;
    }

    /**
     * 解析excel
     * @param inputStream
     * @param filename
     * @throws Exception
     */
    public void importExcel(InputStream inputStream, String filename) throws Exception {

        workbook = getWorkbook(inputStream, filename);
        selectSheet = workbook.getSheet(excelConfig.sheetName());
        // 默认第一个
        if (selectSheet == null) {
            selectSheet = workbook.getSheetAt(0);
        }
    }

    public void setSheet(String name) {
        selectSheet = workbook.getSheet(name);
    }

    public void setSheet(int index) {
        selectSheet = workbook.getSheetAt(index);
    }

    /**
     * 批量导出数据对象， 不符合条件就会停 抛出异常
     * @return
     * @throws Exception ExcelNullpointExcetion
     */
    public List<T> getAllList() throws ExcelNullpointExcetion, ExcelChanelException {
        Sheet sheet = selectSheet;
        if (sheet == null) {
            throw  new RuntimeException("sheet未选择");
        }

        Row row = sheet.getRow(sheet.getFirstRowNum());

        Map<Integer, Excelmode> indexmap = new HashMap<>();
        for (Cell cell: row) {
            if (cell == null || stringEmpty(cell.toString().trim())) {
                continue;
            }

            for (Excelmode excelmode: excelmodeList) {

                if (cell.toString().equals(excelmode.name)) {
                    indexmap.put(cell.getColumnIndex(), excelmode);
                }
            }
        }

        List<T> objs = new ArrayList<>();

        for (Row rown : sheet) {
            if (rown.getRowNum() == sheet.getFirstRowNum()) {
                continue;
            }
            T obj = null;
            try {
                obj = cls.newInstance();
            } catch (InstantiationException | IllegalAccessException e) {
                e.printStackTrace();
            }

            if (stringEmpty(rown.getCell(1).toString())){
                break;
            }

            for (Cell cell: rown) {

                Excelmode excelmode = indexmap.get(cell.getColumnIndex());

                if (excelmode == null) {
                    continue;
                }

                try {
                    excelmode.setmethod.invoke(obj, getCellValue(excelmode.annotation, cell, excelmode.fieldcls));
                } catch (ExcelNullpointExcetion e){
                    throw e;
                } catch (Exception e) {
                    e.fillInStackTrace();
                    throw new ExcelChanelException(cell.getRowIndex(), cell.getColumnIndex());
                }

            }
            objs.add(obj);
        }
        return objs;
    }

    /**
     * 获取当前行对象
     * @param index
     * @return
     * @throws ExcelChanelException
     * @throws ExcelNullpointExcetion
     */
    public T getObjByRow(int index) throws ExcelChanelException, ExcelNullpointExcetion {

        Sheet sheet = selectSheet;
        if (sheet == null) {
            throw  new RuntimeException("sheet未选择");
        }
        T obj = null;
        try {
            obj = cls.newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
            return obj;
        }
        Row rown = sheet.getRow(index);

        for (Excelmode excelmode: excelmodeList) {

            Cell cell = rown.getCell(excelmode.index);
            try {

                excelmode.setmethod.invoke(obj, getCellValue(excelmode.annotation, cell, excelmode.fieldcls));
            } catch (ExcelNullpointExcetion e){
                throw e;
            } catch (Exception e) {
                throw new ExcelChanelException(cell.getRowIndex(), cell.getColumnIndex());
            }
        }
        return obj;
    }

    public int getLastIndex() {
        return selectSheet.getLastRowNum();
    }

    /**
     * 获取最后一行的index
     * 注 ：最后一个行默认为所有字段都不存在值的行
     * @return
     */
    public int getNotNullLastIndex() {

        if (lastIndex == -1) {
            Sheet sheet = selectSheet;
            start:for (Row rown : sheet) {

                for (Excelmode excelmode: excelmodeList) {

                    Cell cell = rown.getCell(excelmode.index);
                    if (cell != null && !stringEmpty(cell.toString())) {
                        continue start;
                    }
                }
                lastIndex = rown.getRowNum();
                break;
            }
            if (lastIndex == -1) {
                lastIndex = sheet.getLastRowNum();
            }
        }
        return lastIndex;
    }

    private Object getCellValue(Annotation annotation, Cell cell, Class cls) throws ExcelNullpointExcetion, ParseException {

        String value = cell.toString();

        if (annotation.annotationType().equals(CellDate.class)) {
           CellDate ac = (CellDate)annotation;
           if (stringEmpty(value)) {
               if (ac.isMust()) {
                   throw new ExcelNullpointExcetion(cell.getRowIndex(), cell.getColumnIndex());
               }
                return null;
           }
           if (!stringEmpty(ac.prefix())) {
               value = value.substring(ac.prefix().length());
           }
           if (!stringEmpty(ac.suffix())) {
               value = value.substring(0, value.length()-ac.suffix().length());
           }

           SimpleDateFormat f = new SimpleDateFormat(ac.formatStr());

           return f.parse(value);
        }

        if (annotation.annotationType().equals(CellBoolean.class)) {
            CellBoolean ac = (CellBoolean)annotation;
            if (stringEmpty(value)) {
                if (ac.isMust()) {
                    throw new ExcelNullpointExcetion(cell.getRowIndex(), cell.getColumnIndex());
                }
                return null;
            }
            if (!stringEmpty(ac.prefix())) {
                value = value.substring(ac.prefix().length());
            }
            if (!stringEmpty(ac.suffix())) {
                value = value.substring(0, value.length()-ac.suffix().length());
            }

            if (value.equals(ac.falseValue())) {
                return false;
            }

            if (value.equals(ac.tureValue())) {
                return true;
            }

            return null;
        }

        if (annotation.annotationType().equals(CellDouble.class)) {
            CellDouble ac = (CellDouble)annotation;
            if (stringEmpty(value)) {
                if (ac.isMust()) {
                    throw new ExcelNullpointExcetion(cell.getRowIndex(), cell.getColumnIndex());
                }
                return null;
            }
            if (!stringEmpty(ac.prefix())) {
                value = value.substring(ac.prefix().length());
            }
            if (!stringEmpty(ac.suffix())) {
                value = value.substring(0, value.length()-ac.suffix().length());
            }

            if (cls.equals(Double.class)) {
                return Double.parseDouble(value);
            }

            if (cls.equals(Float.class)) {
                return Float.parseFloat(value);
            }
        }

        if (annotation.annotationType().equals(CellSelect.class)) {
            CellSelect ac = (CellSelect)annotation;
            if (!stringEmpty(ac.prefix())) {
                value = value.substring(ac.prefix().length());
            }
            if (!stringEmpty(ac.suffix())) {
                value = value.substring(0, value.length()-ac.suffix().length());
            }

            for (int i = 0;i < ac.values().length; i++) {
                if (value.equals(ac.values()[i])) {
                    value = ac.keys()[i];
                    break;
                }
            }

            if (cls.equals(Integer.class)) {
                return Integer.parseInt(value);
            }

            return value;
        }

        if (annotation.annotationType().equals(CellValue.class)) {
            CellValue ac = (CellValue)annotation;
            if (!stringEmpty(ac.prefix())) {
                value = value.substring(ac.prefix().length());
            }
            if (!stringEmpty(ac.suffix())) {
                value = value.substring(0, value.length()-ac.suffix().length());
            }

            if (cls.equals(String.class)) {
                return value;
            }

            if (cls.equals(Integer.class)) {
                return Integer.parseInt(value);
            }

            if (cls.equals(Double.class)) {
                return Double.parseDouble(value);
            }

            if (cls.equals(Boolean.class)) {
                return Boolean.parseBoolean(value);
            }
        }

        return null;
    }

    /**
     * 描述：根据文件后缀，自适应上传文件的版本
     * @param inStr,fileName
     * @return
     * @throws Exception
     */
    public  Workbook getWorkbook(InputStream inStr,String fileName) throws Exception{
        Workbook wb = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if(ExcelEnum.V2003.getSuffix().equals(fileType)){
            wb = new HSSFWorkbook(inStr);  //2003-
        }else if(ExcelEnum.V2007.getSuffix().equals(fileType)){
            wb = new XSSFWorkbook(inStr);  //2007
        }else{
            throw new Exception("解析的文件格式有误！");
        }
        return wb;
    }

    private List<Excelmode> handleCls(Class cls) throws NoSuchMethodException {

        List<Excelmode> list = new ArrayList<>();

        Field[] fields = cls.getDeclaredFields();

        for (Field field: fields) {
            Excelmode excelmode = new Excelmode();
            CellValue cellValue = field.getAnnotation(CellValue.class);
            if (cellValue != null) {
                excelmode.name = cellValue.name();
                excelmode.annotation = cellValue;
                excelmode.index = cellValue.index();
                excelmode.fieldcls = field.getType();
                excelmode.method = cls.getMethod(getFieldString(field));
                excelmode.setmethod = cls.getMethod(setFieldString(field), field.getType());
                list.add(excelmode);
                continue;
            }

            CellSelect cellSelect = field.getAnnotation(CellSelect.class);
            if (cellSelect != null) {
                excelmode.name = cellSelect.name();
                excelmode.annotation = cellSelect;
                excelmode.index = cellSelect.index();
                excelmode.fieldcls = field.getType();
                excelmode.method = cls.getMethod(getFieldString(field));
                excelmode.setmethod = cls.getMethod(setFieldString(field), field.getType());
                list.add(excelmode);
                continue;
            }

            CellDouble cellDouble = field.getAnnotation(CellDouble.class);
            if (cellDouble != null) {
                excelmode.name = cellDouble.name();
                excelmode.annotation = cellDouble;
                excelmode.index = cellDouble.index();
                excelmode.fieldcls = field.getType();
                excelmode.method = cls.getMethod(getFieldString(field));
                excelmode.setmethod = cls.getMethod(setFieldString(field), field.getType());
                list.add(excelmode);
                continue;
            }

            CellDate cellDate = field.getAnnotation(CellDate.class);
            if (cellDate != null) {
                excelmode.name = cellDate.name();
                excelmode.annotation = cellDate;
                excelmode.index = cellDate.index();
                excelmode.fieldcls = field.getType();
                excelmode.method = cls.getMethod(getFieldString(field));
                excelmode.setmethod = cls.getMethod(setFieldString(field), field.getType());
                list.add(excelmode);
                continue;
            }

            CellBoolean cellBoolean = field.getAnnotation(CellBoolean.class);
            if (cellBoolean != null) {
                excelmode.name = cellBoolean.name();
                excelmode.annotation = cellBoolean;
                excelmode.index = cellBoolean.index();
                excelmode.fieldcls = field.getType();
                excelmode.method = cls.getMethod(getFieldString(field));
                excelmode.setmethod = cls.getMethod(setFieldString(field), field.getType());
                list.add(excelmode);
                continue;
            }
        }

        return list.stream().sorted(Comparator.comparingInt(o -> o.index)).collect(Collectors.toList());
    }

    private Workbook getWorkbook(ExcelEnum excelVersion) {
        switch (excelVersion) {
            case V2003: return new XSSFWorkbook();
            case V2007: return new SXSSFWorkbook();
            default: return null;
        }
    }

    private String getFieldString(Field field){

        String name = field.getName();
        char c = name.charAt(0);
        if (c >= 'a' && c <= 'z') {
            c = (char) (c - 32);
        }

        return "get" + c + name.substring(1);
    }

    private String setFieldString(Field field){

        String name = field.getName();
        char c = name.charAt(0);
        if (c >= 'a' && c <= 'z') {
            c = (char) (c - 32);
        }

        return "set" + c + name.substring(1);
    }

    private boolean stringEmpty(String str) {

        return str == null || "".equals(str);
    }

    private static String cellValue(CellValue cellValue, Object obj) {

        if (obj == null) {
            return "";
        }

        String value = "";

        String prefix = cellValue.prefix();
        String suffix = cellValue.suffix();

        value = obj.toString();

        return prefix + value + suffix;
    }


    private String cellSelect(CellSelect cellSelect, Object obj) {

        if (obj == null) {
            return "";
        }

        String value = obj.toString();

        String prefix = cellSelect.prefix();
        String suffix = cellSelect.suffix();

        String[] keys = cellSelect.keys();
        String[] values = cellSelect.values();

        for (int i = 0;i < keys.length; i++) {
            if (keys[i].equals(value)) {
                value = values[i];
                break;
            }
        }
        return prefix + value + suffix;
    }

    private String cellDouble(CellDouble cellDouble, Object obj) {

        if (obj == null) {
            return "";
        }

        String value = obj.toString();

        String prefix = cellDouble.prefix();
        String suffix = cellDouble.suffix();

        int fied = cellDouble.fixed();

        if (obj instanceof Double || obj instanceof Float) {
            value = String.format("%."+fied+"f", obj);
        }

        return prefix + value + suffix;
    }

    private String cellDate(CellDate cellDate, Object obj) {

        if (obj == null) {
            return "";
        }

        String value = obj.toString();

        String prefix = cellDate.prefix();
        String suffix = cellDate.suffix();

        String format = cellDate.formatStr();

        if (obj instanceof Date) {
            SimpleDateFormat f = new SimpleDateFormat(format);
            value = f.format((Date)obj);
        }

        return prefix + value + suffix;
    }

    private String cellBoolean(CellBoolean cellBoolean, Object obj) {

        if (obj == null) {
            return "";
        }

        String value = obj.toString();

        String prefix = cellBoolean.prefix();
        String suffix = cellBoolean.suffix();

        String falseValue = cellBoolean.falseValue();
        String trueValue = cellBoolean.tureValue();

        if (obj instanceof Boolean) {

            value = (Boolean)obj?trueValue:falseValue;
        }

        return prefix + value + suffix;
    }

    class Excelmode{

        String name;

        Method method;

        Method setmethod;

        Class fieldcls;

        int index;

        Annotation annotation;

        @Override
        public String toString() {
            return "Excelmode{" +
                    "name='" + name + '\'' +
                    ", index=" + index +
                    '}';
        }
    }
}
