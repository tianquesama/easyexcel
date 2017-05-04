package com.github.tq.easyexcel.annotation;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

import org.apache.poi.ss.usermodel.CellType;

/**
 * Created by nijun on 2017/4/18.
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface Cell {

    int order() default 0;

    CellType type() default CellType.BLANK;

    String title() default "";

    String dateFormatPattern() default "yyyy/MM/dd HH:mm:ss";

    String errorValue() default "";
}
