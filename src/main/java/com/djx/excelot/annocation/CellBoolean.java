package com.djx.excelot.annocation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface CellBoolean {

    String name();

    int index();

    boolean isMust() default false;

    String tureValue() default "是";

    String falseValue() default "否";

    String prefix() default "";

    String suffix() default "";
}
