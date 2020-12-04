package com.djx.excelot.annocation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface CellDate {

    String formatStr();

    String name();

    int index();

    boolean isMust() default false;

    String prefix() default "";

    String suffix() default "";
}
