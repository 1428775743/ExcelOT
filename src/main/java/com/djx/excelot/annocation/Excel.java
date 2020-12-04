package com.djx.excelot.annocation;

import com.djx.excelot.constants.ExcelEnum;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
public @interface Excel {

    String name();

    ExcelEnum version() default ExcelEnum.V2007;

    String sheetName();
}
