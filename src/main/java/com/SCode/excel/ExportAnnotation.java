package com.SCode.excel;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

import org.apache.poi.ss.usermodel.Cell;

@Retention(RetentionPolicy.RUNTIME)
public @interface ExportAnnotation {

    String name() default "";

    int tpye() default Cell.CELL_TYPE_STRING;
    
    int sort() default 0;
    
}
