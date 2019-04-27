package com.balance.excel.merge.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelStyle {
    String dataFormat() default "";
    short fillForegroundColor() default 0;
    FillPatternType fillPattern() default FillPatternType.NO_FILL;
    HorizontalAlignment alignment() default HorizontalAlignment.GENERAL;
    BorderStyle borderStyle() default BorderStyle.NONE;
}
