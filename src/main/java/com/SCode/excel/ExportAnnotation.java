package com.SCode.excel;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

import org.apache.poi.ss.usermodel.Cell;
/**
 * 
 * Excel 导出注解
 * @author  shizj
 * @version  [1.0, 2016年5月17日]
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface ExportAnnotation {
    /**
     * 表头名称
     * @return
     */
    String name() default "";
    /**
     * 单元格类型
     * @return
     */
    int tpye() default Cell.CELL_TYPE_STRING;
    /**
     * 排序
     * @return
     */
    int sort() default 0;
    
}
